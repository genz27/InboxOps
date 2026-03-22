from __future__ import annotations

import base64
import csv
import hmac
import json
import os
import re
from dataclasses import asdict, is_dataclass
from functools import wraps
from io import StringIO
from pathlib import Path
from typing import Any, Callable

from flask import Flask, jsonify, render_template, request, session

from services.outlook_manager import (
    FlagStateUpdateRequest,
    MailboxConfig,
    MailboxDetailRequest,
    MailboxError,
    MailboxManager,
    MailboxQuery,
    MessageDeleteRequest,
    MessageMoveRequest,
    ReadStateUpdateRequest,
)
from services.storage import MailboxProfile, MailboxStore, MailboxStoreError

MAILBOX_IMPORT_DELIMITER = "----"
MAILBOX_IMPORT_TABULAR_DELIMITERS = ",\t;|"
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
CLIENT_ID_PATTERN = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.IGNORECASE)


def create_app(
    manager: MailboxManager | None = None,
    store: MailboxStore | None = None,
    *,
    database_path: str | Path | None = None,
    admin_username: str | None = None,
    admin_password: str | None = None,
    public_api_key: str | None = None,
) -> Flask:
    app = Flask(__name__)
    app.config["JSON_AS_ASCII"] = False
    app.config["SECRET_KEY"] = os.getenv("MAIL_ADMIN_SECRET_KEY", "change-me-before-production")

    mailbox_manager = manager or MailboxManager()
    mailbox_store = store or MailboxStore(database_path or os.getenv("MAILBOX_DB_PATH", "data/mailboxes.db"))
    admin_user = admin_username or os.getenv("MAIL_ADMIN_USERNAME", "admin")
    admin_pass = admin_password or os.getenv("MAIL_ADMIN_PASSWORD", "admin123456")
    access_key = public_api_key if public_api_key is not None else os.getenv("INBOXOPS_API_KEY", "")

    def auth_required(handler: Callable[..., Any]) -> Callable[..., Any]:
        @wraps(handler)
        def wrapper(*args: Any, **kwargs: Any) -> Any:
            if not session.get("admin_authenticated"):
                raise MailboxError("请先登录管理员账号", code="unauthorized", status_code=401)
            return handler(*args, **kwargs)

        return wrapper

    def api_key_required(handler: Callable[..., Any]) -> Callable[..., Any]:
        @wraps(handler)
        def wrapper(*args: Any, **kwargs: Any) -> Any:
            if not access_key:
                raise MailboxError("项目未配置访问 Key", code="api_key_not_configured", status_code=503)

            provided_key = _extract_access_key_from_request()
            if not provided_key or not hmac.compare_digest(provided_key, access_key):
                raise MailboxError("访问 Key 无效", code="invalid_api_key", status_code=401)
            return handler(*args, **kwargs)

        return wrapper

    @app.get("/")
    def index() -> str:
        return render_template("index.html")

    @app.get("/favicon.ico")
    def favicon() -> tuple[str, int]:
        return "", 204

    @app.get("/api/health")
    def health() -> Any:
        return jsonify({"status": "ok", "service": "inboxops"})

    @app.get("/api/auth/me")
    def auth_me() -> Any:
        authenticated = bool(session.get("admin_authenticated"))
        return jsonify(
            {
                "authenticated": authenticated,
                "username": session.get("admin_username") if authenticated else None,
            }
        )

    @app.post("/api/auth/login")
    def auth_login() -> Any:
        payload = request.get_json(silent=True) or {}
        username = _require_text(payload, "username", "管理员账号不能为空")
        password = _require_text(payload, "password", "管理员密码不能为空")

        if not hmac.compare_digest(username, admin_user) or not hmac.compare_digest(password, admin_pass):
            raise MailboxError("管理员账号或密码错误", code="invalid_credentials", status_code=401)

        session["admin_authenticated"] = True
        session["admin_username"] = admin_user
        return jsonify({"authenticated": True, "username": admin_user})

    @app.post("/api/auth/logout")
    @auth_required
    def auth_logout() -> Any:
        session.clear()
        return jsonify({"authenticated": False})

    @app.get("/api/mailboxes")
    @auth_required
    def list_mailboxes() -> Any:
        query = _optional_text(request.args.get("q")) or ""
        page = _parse_positive_int(request.args.get("page"), field_name="page", default=1, minimum=1)
        page_size = _parse_positive_int(
            request.args.get("page_size"),
            field_name="page_size",
            default=20,
            minimum=1,
            maximum=100,
        )
        items, total = mailbox_store.search_mailboxes_summary(query, page=page, page_size=page_size)
        total_pages = (total + page_size - 1) // page_size if total else 0
        return jsonify(
            {
                "items": _to_jsonable(items),
                "meta": {
                    "q": query,
                    "page": page,
                    "page_size": page_size,
                    "total": total,
                    "total_pages": total_pages,
                    "has_prev": page > 1,
                    "has_next": page < total_pages,
                },
            }
        )

    @app.post("/api/mailboxes")
    @auth_required
    def create_mailbox() -> Any:
        payload = request.get_json(silent=True) or {}
        mailbox = mailbox_store.create_mailbox(_extract_mailbox_payload(payload))
        return jsonify({"mailbox": _to_jsonable(mailbox)}), 201

    @app.post("/api/mailboxes/import")
    @auth_required
    def import_mailboxes() -> Any:
        payload = request.get_json(silent=True) or {}
        raw_text = _require_text(payload, "raw_text", "批量导入文本不能为空")
        preferred_method = _normalize_method(payload.get("preferred_method") or "graph_api")
        parsed_payloads = _parse_import_mailboxes(
            raw_text,
            preferred_method=preferred_method,
            allow_missing_email=True,
        )
        hydrated_payloads = _hydrate_import_mailboxes_missing_email(
            mailbox_manager,
            parsed_payloads,
            preferred_method=preferred_method,
        )
        summary, mailboxes = mailbox_store.import_mailboxes(
            hydrated_payloads
        )
        return jsonify({"summary": summary, "mailboxes": _to_jsonable(mailboxes)})

    @app.get("/api/mailboxes/<int:mailbox_id>")
    @auth_required
    def get_mailbox(mailbox_id: int) -> Any:
        mailbox = _get_mailbox_or_404(mailbox_store, mailbox_id)
        return jsonify({"mailbox": _to_jsonable(mailbox)})

    @app.put("/api/mailboxes/<int:mailbox_id>")
    @auth_required
    def update_mailbox(mailbox_id: int) -> Any:
        payload = request.get_json(silent=True) or {}
        mailbox = mailbox_store.update_mailbox(mailbox_id, _extract_mailbox_payload(payload, partial=True))
        return jsonify({"mailbox": _to_jsonable(mailbox)})

    @app.delete("/api/mailboxes/<int:mailbox_id>")
    @auth_required
    def delete_mailbox(mailbox_id: int) -> Any:
        deleted = mailbox_store.delete_mailbox(mailbox_id)
        if not deleted:
            raise MailboxError("邮箱档案不存在", code="mailbox_not_found", status_code=404)
        return jsonify({"deleted": True, "mailbox_id": mailbox_id})

    @app.post("/api/mailboxes/delete/batch")
    @auth_required
    def batch_delete_mailboxes() -> Any:
        payload = request.get_json(silent=True) or {}
        mailbox_ids = _parse_mailbox_ids(payload.get("mailbox_ids"), maximum=100)
        results: list[dict[str, Any]] = []
        succeeded = 0
        failed = 0

        for mailbox_id in mailbox_ids:
            profile = mailbox_store.get_mailbox(mailbox_id)
            if not profile:
                results.append(
                    {
                        "mailbox_id": mailbox_id,
                        "label": "",
                        "email": "",
                        "success": False,
                        "message": "邮箱档案不存在",
                    }
                )
                failed += 1
                continue

            deleted = mailbox_store.delete_mailbox(mailbox_id)
            if deleted:
                results.append(
                    {
                        "mailbox_id": mailbox_id,
                        "label": profile.label,
                        "email": profile.email,
                        "success": True,
                        "message": "邮箱档案已删除",
                    }
                )
                succeeded += 1
            else:
                results.append(
                    {
                        "mailbox_id": mailbox_id,
                        "label": profile.label,
                        "email": profile.email,
                        "success": False,
                        "message": "邮箱档案删除失败",
                    }
                )
                failed += 1

        return jsonify(
            {
                "results": results,
                "summary": {
                    "processed": len(results),
                    "succeeded": succeeded,
                    "failed": failed,
                },
            }
        )

    @app.post("/api/mailboxes/test-connection")
    @auth_required
    def test_mailbox_connection() -> Any:
        payload = request.get_json(silent=True) or {}
        config, method = _build_runtime_config_from_payload(payload)
        message = _probe_mailbox_connection(mailbox_manager, config=config, method=method)
        return jsonify(
            {
                "success": True,
                "method": method,
                "label": _label_for_method(method),
                "message": message,
            }
        )

    @app.post("/api/mailboxes/test-connection/batch")
    @auth_required
    def test_mailbox_connection_batch() -> Any:
        payload = request.get_json(silent=True) or {}
        mailbox_ids = _parse_mailbox_ids(payload.get("mailbox_ids"), maximum=50)
        raw_method = payload.get("method")
        forced_method = _normalize_method(raw_method) if raw_method is not None else None
        results: list[dict[str, Any]] = []
        succeeded = 0
        failed = 0

        for raw_mailbox_id in mailbox_ids:
            if not isinstance(raw_mailbox_id, int):
                raise MailboxError("mailbox_ids 中的每一项都必须是整数", code="invalid_mailbox_ids")

            profile = mailbox_store.get_mailbox(raw_mailbox_id)
            if not profile:
                results.append(
                    {
                        "mailbox_id": raw_mailbox_id,
                        "label": "",
                        "email": "",
                        "method": forced_method or "",
                        "success": False,
                        "message": "邮箱档案不存在",
                    }
                )
                failed += 1
                continue

            method = forced_method or profile.preferred_method
            try:
                message = _probe_mailbox_connection(
                    mailbox_manager,
                    config=_profile_to_config(profile, method=method),
                    method=method,
                )
                results.append(
                    {
                        "mailbox_id": profile.id,
                        "label": profile.label,
                        "email": profile.email,
                        "method": method,
                        "success": True,
                        "message": message,
                    }
                )
                succeeded += 1
            except Exception as exc:  # noqa: BLE001
                results.append(
                    {
                        "mailbox_id": profile.id,
                        "label": profile.label,
                        "email": profile.email,
                        "method": method,
                        "success": False,
                        "message": str(exc),
                    }
                )
                failed += 1

        return jsonify(
            {
                "results": results,
                "summary": {
                    "processed": len(results),
                    "succeeded": succeeded,
                    "failed": failed,
                },
            }
        )

    @app.post("/api/mailboxes/preferred-method/batch")
    @auth_required
    def batch_update_preferred_method() -> Any:
        payload = request.get_json(silent=True) or {}
        mailbox_ids = _parse_mailbox_ids(payload.get("mailbox_ids"), maximum=100)
        preferred_method = _normalize_method(payload.get("preferred_method"))
        results: list[dict[str, Any]] = []
        succeeded = 0
        failed = 0

        for mailbox_id in mailbox_ids:
            profile = mailbox_store.get_mailbox(mailbox_id)
            if not profile:
                results.append(
                    {
                        "mailbox_id": mailbox_id,
                        "label": "",
                        "email": "",
                        "preferred_method": preferred_method,
                        "success": False,
                        "message": "邮箱档案不存在",
                    }
                )
                failed += 1
                continue

            updated = mailbox_store.update_mailbox(
                mailbox_id,
                {
                    "preferred_method": preferred_method,
                },
            )
            results.append(
                {
                    "mailbox_id": updated.id,
                    "label": updated.label,
                    "email": updated.email,
                    "preferred_method": updated.preferred_method,
                    "success": True,
                    "message": f"默认方法已切换为 {_label_for_method(updated.preferred_method)}",
                }
            )
            succeeded += 1

        return jsonify(
            {
                "results": results,
                "summary": {
                    "processed": len(results),
                    "succeeded": succeeded,
                    "failed": failed,
                },
            }
        )

    @app.post("/api/mailbox/overview")
    @auth_required
    def mailbox_overview() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        config = _profile_to_config(profile)
        overview = mailbox_manager.get_overview(config)
        return jsonify({"mailbox": _to_jsonable(profile), "overview": _to_jsonable(overview)})

    @app.post("/api/mailbox/folders")
    @auth_required
    def mailbox_folders() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        config = _profile_to_config(profile, method=method)
        folders = mailbox_manager.list_folders(config, method)
        mailbox_store.cache_folders(profile.id, method, _to_jsonable(folders))
        return jsonify({"mailbox": _to_jsonable(profile), "method": method, "folders": _to_jsonable(folders)})

    @app.post("/api/mailbox/messages")
    @auth_required
    def mailbox_messages() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        query = _build_query(payload, default_method=profile.preferred_method)
        config = _profile_to_config(profile, method=query.method)
        result = mailbox_manager.list_messages(config, query)
        messages = result.messages if hasattr(result, "messages") else result
        count = getattr(result, "returned", len(messages))
        mailbox_store.cache_messages(profile.id, query.method, _to_jsonable(messages))
        return jsonify(
            {
                "mailbox": _to_jsonable(profile),
                "method": query.method,
                "folder": query.folder,
                "messages": _to_jsonable(messages),
                "count": count,
                "meta": {
                    "total": getattr(result, "total", count),
                    "returned": count,
                    "page": getattr(result, "page", query.page),
                    "page_size": getattr(result, "page_size", query.page_size),
                    "total_pages": getattr(result, "total_pages", 1 if count else 0),
                    "has_prev": getattr(result, "has_prev", False),
                    "has_next": getattr(result, "has_next", False),
                    "folder": getattr(result, "folder", query.folder),
                },
            }
        )

    @app.post("/api/key/mailbox/messages")
    @api_key_required
    def mailbox_messages_by_key() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox_by_email(mailbox_store, payload)
        query = _build_query(payload, default_method=profile.preferred_method)
        config = _profile_to_config(profile, method=query.method)
        result = mailbox_manager.list_messages(config, query)
        messages = result.messages if hasattr(result, "messages") else result
        count = getattr(result, "returned", len(messages))
        mailbox_store.cache_messages(profile.id, query.method, _to_jsonable(messages))
        return jsonify(
            {
                "mailbox": _profile_to_public_mailbox(profile),
                "method": query.method,
                "folder": query.folder,
                "messages": _to_jsonable(messages),
                "count": count,
                "meta": {
                    "total": getattr(result, "total", count),
                    "returned": count,
                    "page": getattr(result, "page", query.page),
                    "page_size": getattr(result, "page_size", query.page_size),
                    "total_pages": getattr(result, "total_pages", 1 if count else 0),
                    "has_prev": getattr(result, "has_prev", False),
                    "has_next": getattr(result, "has_next", False),
                    "folder": getattr(result, "folder", query.folder),
                },
            }
        )

    @app.post("/api/mailbox/message")
    @auth_required
    def mailbox_message() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        detail_request = _build_detail_request(payload, default_method=profile.preferred_method)
        config = _profile_to_config(profile, method=detail_request.method)
        message = mailbox_manager.get_message_detail(config, detail_request)
        mailbox_store.cache_message(profile.id, detail_request.method, _to_jsonable(message))
        return jsonify({"mailbox": _to_jsonable(profile), "message": _to_jsonable(message)})

    @app.post("/api/key/mailbox/message")
    @api_key_required
    def mailbox_message_by_key() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox_by_email(mailbox_store, payload)
        detail_request = _build_detail_request(payload, default_method=profile.preferred_method)
        config = _profile_to_config(profile, method=detail_request.method)
        message = mailbox_manager.get_message_detail(config, detail_request)
        mailbox_store.cache_message(profile.id, detail_request.method, _to_jsonable(message))
        return jsonify(
            {
                "mailbox": _profile_to_public_mailbox(profile),
                "message": _to_jsonable(message),
            }
        )

    @app.post("/api/mailbox/message/read-state")
    @auth_required
    def mailbox_message_read_state() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        state_request = _build_read_state_request(payload, default_method=profile.preferred_method)
        config = _profile_to_config(profile, method=state_request.method)
        mailbox_manager.update_read_state(config, state_request)
        message = mailbox_manager.get_message_detail(
            config,
            MailboxDetailRequest(
                method=state_request.method,
                message_id=state_request.message_id,
                folder=state_request.folder,
            ),
        )
        mailbox_store.cache_message(profile.id, state_request.method, _to_jsonable(message))
        return jsonify({"mailbox": _to_jsonable(profile), "message": _to_jsonable(message)})

    @app.post("/api/mailbox/message/flag-state")
    @auth_required
    def mailbox_message_flag_state() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        state_request = _build_flag_state_request(payload, default_method=profile.preferred_method)
        config = _profile_to_config(profile, method=state_request.method)
        mailbox_manager.update_flag_state(config, state_request)
        message = mailbox_manager.get_message_detail(
            config,
            MailboxDetailRequest(
                method=state_request.method,
                message_id=state_request.message_id,
                folder=state_request.folder,
            ),
        )
        mailbox_store.cache_message(profile.id, state_request.method, _to_jsonable(message))
        return jsonify({"mailbox": _to_jsonable(profile), "message": _to_jsonable(message)})

    @app.post("/api/mailbox/message/move")
    @auth_required
    def mailbox_message_move() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        move_request = _build_move_request(payload, default_method=profile.preferred_method)
        config = _profile_to_config(profile, method=move_request.method)
        result = mailbox_manager.move_message(config, move_request)
        mailbox_store.update_cached_message_state(
            profile.id,
            move_request.method,
            move_request.message_id,
            folder_id=move_request.destination_folder,
        )
        return jsonify({"mailbox": _to_jsonable(profile), "result": _to_jsonable(result)})

    @app.post("/api/mailbox/message/delete")
    @auth_required
    def mailbox_message_delete() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        delete_request = _build_delete_request(payload, default_method=profile.preferred_method)
        config = _profile_to_config(profile, method=delete_request.method)
        result = mailbox_manager.delete_message(config, delete_request)
        mailbox_store.remove_cached_message(profile.id, delete_request.method, delete_request.message_id)
        return jsonify({"mailbox": _to_jsonable(profile), "result": _to_jsonable(result)})

    @app.post("/api/mailbox/messages/actions/batch")
    @auth_required
    def mailbox_messages_batch_actions() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        folder = _optional_text(payload.get("folder")) or "INBOX"
        action = _normalize_message_action(payload.get("action"))
        message_ids = _parse_message_ids(payload.get("message_ids"), maximum=100)
        destination_folder = _optional_text(payload.get("destination_folder"))
        if action == "move" and not destination_folder:
            raise MailboxError("移动邮件时必须提供目标文件夹", code="invalid_destination_folder")
        if action == "archive" and not destination_folder:
            destination_folder = "archive"

        config = _profile_to_config(profile, method=method)
        results: list[dict[str, Any]] = []
        succeeded = 0
        failed = 0

        for message_id in message_ids:
            try:
                if action == "mark_read":
                    result = mailbox_manager.update_read_state(
                        config,
                        ReadStateUpdateRequest(method=method, message_id=message_id, is_read=True, folder=folder),
                    )
                elif action == "mark_unread":
                    result = mailbox_manager.update_read_state(
                        config,
                        ReadStateUpdateRequest(method=method, message_id=message_id, is_read=False, folder=folder),
                    )
                elif action == "flag":
                    result = mailbox_manager.update_flag_state(
                        config,
                        FlagStateUpdateRequest(method=method, message_id=message_id, is_flagged=True, folder=folder),
                    )
                elif action == "unflag":
                    result = mailbox_manager.update_flag_state(
                        config,
                        FlagStateUpdateRequest(method=method, message_id=message_id, is_flagged=False, folder=folder),
                    )
                elif action in {"move", "archive"}:
                    result = mailbox_manager.move_message(
                        config,
                        MessageMoveRequest(
                            method=method,
                            message_id=message_id,
                            destination_folder=destination_folder or "archive",
                            folder=folder,
                        ),
                    )
                else:
                    result = mailbox_manager.delete_message(
                        config,
                        MessageDeleteRequest(method=method, message_id=message_id, folder=folder),
                    )

                results.append(
                    {
                        "message_id": message_id,
                        "success": True,
                        "status": getattr(result, "status", "updated"),
                        "result": _to_jsonable(result),
                    }
                )
                succeeded += 1
            except Exception as exc:  # noqa: BLE001
                results.append(
                    {
                        "message_id": message_id,
                        "success": False,
                        "status": "error",
                        "message": str(exc),
                    }
                )
                failed += 1

        return jsonify(
            {
                "mailbox": _to_jsonable(profile),
                "action": action,
                "results": results,
                "summary": {
                    "processed": len(results),
                    "succeeded": succeeded,
                    "failed": failed,
                },
            }
        )

    @app.post("/api/mailbox/message/draft")
    @auth_required
    def mailbox_message_draft() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        config = _profile_to_config(profile, method=method)
        draft = mailbox_manager.save_draft(config, _build_compose_payload(payload, require_to=False))
        mailbox_store.cache_message(profile.id, method, draft)
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="draft_saved",
            target_type="message",
            target_id=str(draft.get("message_id", "")),
            details={"method": method},
        )
        return jsonify({"mailbox": _to_jsonable(profile), "message": _to_jsonable(draft)})

    @app.post("/api/mailbox/message/send")
    @auth_required
    def mailbox_message_send() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        config = _profile_to_config(profile, method=method)
        result = mailbox_manager.send_message(config, _build_compose_payload(payload, require_to=True))
        if result.get("message_id"):
            mailbox_store.cache_message(profile.id, method, result)
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="message_sent",
            target_type="message",
            target_id=str(result.get("message_id", "")),
            details={"method": method},
        )
        return jsonify({"mailbox": _to_jsonable(profile), "message": _to_jsonable(result)})

    @app.post("/api/mailbox/message/reply")
    @auth_required
    def mailbox_message_reply() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        message_id = _require_text(payload, "message_id", "缺少邮件标识")
        config = _profile_to_config(profile, method=method)
        result = mailbox_manager.reply_message(
            config,
            message_id,
            _build_compose_payload(payload, require_to=False),
            reply_all=False,
        )
        mailbox_store.cache_message(profile.id, method, result)
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="message_replied",
            target_type="message",
            target_id=message_id,
            details={"method": method},
        )
        return jsonify({"mailbox": _to_jsonable(profile), "message": _to_jsonable(result)})

    @app.post("/api/mailbox/message/reply-all")
    @auth_required
    def mailbox_message_reply_all() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        message_id = _require_text(payload, "message_id", "缺少邮件标识")
        config = _profile_to_config(profile, method=method)
        result = mailbox_manager.reply_message(
            config,
            message_id,
            _build_compose_payload(payload, require_to=False),
            reply_all=True,
        )
        mailbox_store.cache_message(profile.id, method, result)
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="message_replied_all",
            target_type="message",
            target_id=message_id,
            details={"method": method},
        )
        return jsonify({"mailbox": _to_jsonable(profile), "message": _to_jsonable(result)})

    @app.post("/api/mailbox/message/forward")
    @auth_required
    def mailbox_message_forward() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        message_id = _require_text(payload, "message_id", "缺少邮件标识")
        config = _profile_to_config(profile, method=method)
        result = mailbox_manager.forward_message(
            config,
            message_id,
            _build_compose_payload(payload, require_to=True),
        )
        mailbox_store.cache_message(profile.id, method, result)
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="message_forwarded",
            target_type="message",
            target_id=message_id,
            details={"method": method},
        )
        return jsonify({"mailbox": _to_jsonable(profile), "message": _to_jsonable(result)})

    @app.post("/api/mailbox/message/attachment/upload")
    @auth_required
    def mailbox_message_attachment_upload() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        message_id = _require_text(payload, "message_id", "缺少邮件标识")
        config = _profile_to_config(profile, method=method)
        attachment_payload = _build_attachment_payload(payload)
        attachment = mailbox_manager.upload_attachment(config, message_id, attachment_payload)
        mailbox_store.ensure_cached_message_placeholder(profile.id, method, message_id, folder_id="drafts")
        mailbox_store.upsert_attachment_content(profile.id, method, message_id, attachment)
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="attachment_uploaded",
            target_type="attachment",
            target_id=str(attachment.get("id") or attachment.get("attachment_id", "")),
            details={"message_id": message_id, "method": method},
        )
        return jsonify({"mailbox": _to_jsonable(profile), "attachment": _to_jsonable(attachment)})

    @app.post("/api/mailbox/message/attachment/download")
    @auth_required
    def mailbox_message_attachment_download() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        message_id = _require_text(payload, "message_id", "缺少邮件标识")
        attachment_id = _require_text(payload, "attachment_id", "缺少附件标识")
        config = _profile_to_config(profile, method=method)
        attachment = mailbox_manager.download_attachment(config, message_id, attachment_id)
        mailbox_store.ensure_cached_message_placeholder(
            profile.id,
            method,
            message_id,
            folder_id=_optional_text(payload.get("folder")) or "",
        )
        mailbox_store.upsert_attachment_content(profile.id, method, message_id, attachment)
        return jsonify({"mailbox": _to_jsonable(profile), "attachment": _to_jsonable(attachment)})

    @app.post("/api/mailbox/folder/create")
    @auth_required
    def mailbox_folder_create() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        config = _profile_to_config(profile, method=method)
        folder = mailbox_manager.create_folder(config, payload)
        mailbox_store.cache_folders(profile.id, method, _to_jsonable(mailbox_manager.list_folders(config, method)))
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="folder_created",
            target_type="folder",
            target_id=str(folder.get("id", "")),
            details={"method": method},
        )
        return jsonify({"mailbox": _to_jsonable(profile), "folder": _to_jsonable(folder)})

    @app.post("/api/mailbox/folder/rename")
    @auth_required
    def mailbox_folder_rename() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        config = _profile_to_config(profile, method=method)
        folder = mailbox_manager.rename_folder(config, payload)
        mailbox_store.cache_folders(profile.id, method, _to_jsonable(mailbox_manager.list_folders(config, method)))
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="folder_renamed",
            target_type="folder",
            target_id=str(folder.get("id", "")),
            details={"method": method},
        )
        return jsonify({"mailbox": _to_jsonable(profile), "folder": _to_jsonable(folder)})

    @app.post("/api/mailbox/folder/delete")
    @auth_required
    def mailbox_folder_delete() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        config = _profile_to_config(profile, method=method)
        folder = mailbox_manager.delete_folder(config, payload)
        mailbox_store.cache_folders(profile.id, method, _to_jsonable(mailbox_manager.list_folders(config, method)))
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="folder_deleted",
            target_type="folder",
            target_id=str(folder.get("id", "")),
            details={"method": method},
        )
        return jsonify({"mailbox": _to_jsonable(profile), "folder": _to_jsonable(folder)})

    @app.post("/api/mailbox/message/meta")
    @auth_required
    def mailbox_message_meta() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _normalize_method(payload.get("method") or profile.preferred_method)
        message_id = _require_text(payload, "message_id", "缺少邮件标识")
        if not mailbox_store.get_cached_message(profile.id, method, message_id):
            folder = _optional_text(payload.get("folder")) or "INBOX"
            config = _profile_to_config(profile, method=method)
            try:
                detail = mailbox_manager.get_message_detail(
                    config,
                    MailboxDetailRequest(method=method, message_id=message_id, folder=folder),
                )
                mailbox_store.cache_message(profile.id, method, _to_jsonable(detail))
            except Exception:
                mailbox_store.ensure_cached_message_placeholder(profile.id, method, message_id, folder_id=folder)
        meta = mailbox_store.update_message_meta(
            profile.id,
            method,
            message_id,
            tags=_optional_text_list(payload.get("tags"), field_name="tags"),
            follow_up=_optional_text(payload.get("follow_up")),
            notes=_optional_text(payload.get("notes")),
            snoozed_until=_optional_text(payload.get("snoozed_until")),
            status=_optional_text(payload.get("status")),
        )
        message = mailbox_store.get_cached_message(profile.id, method, message_id)
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="message_meta_updated",
            target_type="message",
            target_id=message_id,
            details={"method": method},
        )
        return jsonify({"mailbox": _to_jsonable(profile), "meta": _to_jsonable(meta), "message": _to_jsonable(message)})

    @app.post("/api/mailboxes/messages/search")
    @auth_required
    def mailbox_messages_search() -> Any:
        payload = request.get_json(silent=True) or {}
        method = _optional_text(payload.get("method"))
        page = _parse_positive_int(payload.get("page"), field_name="page", default=1, minimum=1)
        page_size = _parse_positive_int(payload.get("page_size"), field_name="page_size", default=20, minimum=1, maximum=100)
        mailbox_ids = _parse_optional_mailbox_ids(payload.get("mailbox_ids"))
        tags = _optional_text_list(payload.get("tags"), field_name="tags")
        results, total = mailbox_store.search_messages(
            _optional_text(payload.get("query")) or _optional_text(payload.get("keyword")),
            mailbox_ids=mailbox_ids or None,
            method=_normalize_method(method) if method else None,
            folder=_optional_text(payload.get("folder")),
            tag=tags[0] if tags else _optional_text(payload.get("tag")),
            unread_only=_parse_bool(payload.get("unread_only"), field_name="unread_only", default=False),
            flagged_only=_parse_bool(payload.get("flagged_only"), field_name="flagged_only", default=False),
            has_attachments_only=_parse_bool(
                payload.get("has_attachments_only"),
                field_name="has_attachments_only",
                default=False,
            ),
            include_snoozed=_parse_bool(payload.get("include_snoozed"), field_name="include_snoozed", default=True),
            page=page,
            page_size=page_size,
            sort_order=_normalize_sort_order(payload.get("sort_order")),
        )
        total_pages = (total + page_size - 1) // page_size if total else 0
        return jsonify(
            {
                "items": _to_jsonable(results),
                "meta": {
                    "page": page,
                    "page_size": page_size,
                    "total": total,
                    "total_pages": total_pages,
                    "has_prev": page > 1,
                    "has_next": page < total_pages,
                },
            }
        )

    @app.post("/api/mailbox/thread")
    @auth_required
    def mailbox_thread() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _optional_text(payload.get("method"))
        message_id = _optional_text(payload.get("message_id"))
        conversation_id = _optional_text(payload.get("conversation_id"))
        if message_id and method and not mailbox_store.get_cached_message(profile.id, _normalize_method(method), message_id):
            folder = _optional_text(payload.get("folder")) or "INBOX"
            normalized_method = _normalize_method(method)
            config = _profile_to_config(profile, method=normalized_method)
            try:
                detail = mailbox_manager.get_message_detail(
                    config,
                    MailboxDetailRequest(method=normalized_method, message_id=message_id, folder=folder),
                )
                mailbox_store.cache_message(profile.id, normalized_method, _to_jsonable(detail))
            except Exception:
                pass
        thread = mailbox_store.list_thread_messages(
            profile.id,
            method=_normalize_method(method) if method else None,
            conversation_id=conversation_id,
            message_id=message_id,
        )
        return jsonify({"mailbox": _to_jsonable(profile), "items": _to_jsonable(thread), "count": len(thread)})

    @app.post("/api/mailbox/rules/list")
    @auth_required
    def mailbox_rules_list() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        rules = mailbox_store.list_rules(profile.id, enabled_only=_parse_bool(payload.get("enabled_only"), field_name="enabled_only", default=False))
        return jsonify({"mailbox": _to_jsonable(profile), "items": _to_jsonable(rules), "count": len(rules)})

    @app.post("/api/mailbox/rules")
    @auth_required
    def mailbox_rules_create() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        rule = mailbox_store.create_rule(
            profile.id,
            name=_require_text(payload, "name", "规则名称不能为空"),
            enabled=_parse_bool(payload.get("enabled"), field_name="enabled", default=True),
            priority=_parse_positive_int(payload.get("priority"), field_name="priority", default=100, minimum=1, maximum=9999),
            conditions=_optional_object(payload.get("conditions"), field_name="conditions"),
            actions=_optional_object(payload.get("actions"), field_name="actions"),
        )
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="rule_created",
            target_type="rule",
            target_id=str(rule.id),
            details={"name": rule.name},
        )
        return jsonify({"mailbox": _to_jsonable(profile), "rule": _to_jsonable(rule)}), 201

    @app.post("/api/mailbox/rules/update")
    @auth_required
    def mailbox_rules_update() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        rule_id = _parse_positive_int(payload.get("rule_id"), field_name="rule_id", minimum=1)
        rule = mailbox_store.update_rule(rule_id, profile.id, payload)
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="rule_updated",
            target_type="rule",
            target_id=str(rule.id),
            details={"name": rule.name},
        )
        return jsonify({"mailbox": _to_jsonable(profile), "rule": _to_jsonable(rule)})

    @app.post("/api/mailbox/rules/delete")
    @auth_required
    def mailbox_rules_delete() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        rule_id = _parse_positive_int(payload.get("rule_id"), field_name="rule_id", minimum=1)
        deleted = mailbox_store.delete_rule(rule_id, profile.id)
        if not deleted:
            raise MailboxError("规则不存在", code="rule_not_found", status_code=404)
        mailbox_store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="rule_deleted",
            target_type="rule",
            target_id=str(rule_id),
            details={},
        )
        return jsonify({"deleted": True, "rule_id": rule_id})

    @app.post("/api/mailbox/rules/apply")
    @auth_required
    def mailbox_rules_apply() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox(mailbox_store, payload)
        method = _optional_text(payload.get("method"))
        rules = (
            [mailbox_store.get_rule(_parse_positive_int(payload.get("rule_id"), field_name="rule_id", minimum=1), profile.id)]
            if payload.get("rule_id") is not None
            else mailbox_store.list_rules(profile.id, enabled_only=True)
        )
        active_rules = [rule for rule in rules if rule]
        candidates = mailbox_store.list_cached_messages(
            profile.id,
            method=_normalize_method(method) if method else None,
            folder=_optional_text(payload.get("folder")),
            include_snoozed=True,
            limit=_parse_positive_int(payload.get("limit"), field_name="limit", default=200, minimum=1, maximum=1000),
        )
        applied = _apply_rules_to_messages(
            mailbox_manager,
            mailbox_store,
            profile,
            active_rules,
            candidates,
            method=_normalize_method(method) if method else None,
        )
        return jsonify({"mailbox": _to_jsonable(profile), "results": _to_jsonable(applied), "count": len(applied)})

    @app.post("/api/audit/logs")
    @auth_required
    def audit_logs() -> Any:
        payload = request.get_json(silent=True) or {}
        page = _parse_positive_int(payload.get("page"), field_name="page", default=1, minimum=1)
        page_size = _parse_positive_int(payload.get("page_size"), field_name="page_size", default=50, minimum=1, maximum=100)
        mailbox_id = _parse_optional_single_mailbox_id(payload.get("mailbox_id"))
        items, total = mailbox_store.list_audit_logs(
            mailbox_id=mailbox_id,
            action=_optional_text(payload.get("action")),
            page=page,
            page_size=page_size,
        )
        total_pages = (total + page_size - 1) // page_size if total else 0
        return jsonify(
            {
                "items": _to_jsonable(items),
                "meta": {
                    "page": page,
                    "page_size": page_size,
                    "total": total,
                    "total_pages": total_pages,
                    "has_prev": page > 1,
                    "has_next": page < total_pages,
                },
            }
        )

    @app.post("/api/mailbox/sync/run")
    @auth_required
    def mailbox_sync_run() -> Any:
        payload = request.get_json(silent=True) or {}
        requested_ids = _parse_optional_mailbox_ids(payload.get("mailbox_ids"))
        single_id = _parse_optional_single_mailbox_id(payload.get("mailbox_id"))
        if single_id is not None and single_id not in requested_ids:
            requested_ids.append(single_id)
        profiles = (
            [_get_mailbox_or_404(mailbox_store, mailbox_id) for mailbox_id in requested_ids]
            if requested_ids
            else mailbox_store.list_mailboxes()
        )
        folder_limit = _parse_positive_int(payload.get("folder_limit"), field_name="folder_limit", default=5, minimum=1, maximum=20)
        message_limit = _parse_positive_int(payload.get("message_limit"), field_name="message_limit", default=20, minimum=1, maximum=100)
        include_body = _parse_bool(payload.get("include_body"), field_name="include_body", default=True)
        apply_rules = _parse_bool(payload.get("apply_rules"), field_name="apply_rules", default=True)
        results: list[dict[str, Any]] = []

        for profile in profiles:
            method = _normalize_method(payload.get("method") or profile.preferred_method)
            config = _profile_to_config(profile, method=method)
            scope = {"folder_limit": folder_limit, "message_limit": message_limit, "include_body": include_body}
            job = mailbox_store.create_sync_job(mailbox_id=profile.id, method=method, requested_by="admin", scope=scope)
            processed_messages = 0
            cached_messages = 0
            folders_synced = 0
            error_message = ""
            try:
                folders = mailbox_manager.list_folders(config, method)
                mailbox_store.cache_folders(profile.id, method, _to_jsonable(folders))
                selected_folders = folders[:folder_limit]
                for folder in selected_folders:
                    folder_id = str(getattr(folder, "id", None) or folder.get("id", ""))
                    query = MailboxQuery(
                        method=method,
                        folder=folder_id or "INBOX",
                        top=message_limit,
                        page=1,
                        page_size=message_limit,
                    )
                    result = mailbox_manager.list_messages(config, query)
                    messages = result.messages if hasattr(result, "messages") else result
                    mailbox_store.cache_messages(profile.id, method, _to_jsonable(messages))
                    processed_messages += len(messages)
                    cached_messages += len(messages)
                    folders_synced += 1
                    last_message_at = ""
                    if include_body:
                        for item in messages:
                            message_id = str(getattr(item, "message_id", None) or item.get("message_id", ""))
                            if not message_id:
                                continue
                            detail = mailbox_manager.get_message_detail(
                                config,
                                MailboxDetailRequest(method=method, message_id=message_id, folder=query.folder),
                            )
                            mailbox_store.cache_message(profile.id, method, _to_jsonable(detail))
                            last_message_at = getattr(detail, "received_at", None) or detail.get("received_at", last_message_at)
                    mailbox_store.upsert_sync_state(
                        mailbox_id=profile.id,
                        method=method,
                        folder_id=folder_id or "__mailbox__",
                        cached_messages=len(messages),
                        last_message_at=last_message_at,
                        status="completed",
                    )
                mailbox_store.update_sync_job(
                    job.id,
                    status="completed",
                    processed_messages=processed_messages,
                    cached_messages=cached_messages,
                    folders_synced=folders_synced,
                )
                if apply_rules:
                    _apply_rules_to_messages(
                        mailbox_manager,
                        mailbox_store,
                        profile,
                        mailbox_store.list_rules(profile.id, enabled_only=True),
                        mailbox_store.list_cached_messages(profile.id, method=method, limit=message_limit * folder_limit),
                        method=method,
                    )
                results.append({"mailbox": _to_jsonable(profile), "job": _to_jsonable(mailbox_store.get_sync_job(job.id))})
            except Exception as exc:  # noqa: BLE001
                error_message = str(exc)
                mailbox_store.update_sync_job(
                    job.id,
                    status="failed",
                    processed_messages=processed_messages,
                    cached_messages=cached_messages,
                    folders_synced=folders_synced,
                    error=error_message,
                )
                mailbox_store.record_audit_log(
                    mailbox_id=profile.id,
                    actor="admin",
                    action="sync_failed",
                    target_type="sync_job",
                    target_id=str(job.id),
                    status="failed",
                    details={"error": error_message, "method": method},
                )
                raise

        return jsonify({"results": _to_jsonable(results), "count": len(results)})

    @app.post("/api/mailbox/sync/status")
    @auth_required
    def mailbox_sync_status() -> Any:
        payload = request.get_json(silent=True) or {}
        requested_ids = _parse_optional_mailbox_ids(payload.get("mailbox_ids"))
        single_id = _parse_optional_single_mailbox_id(payload.get("mailbox_id"))
        if single_id is not None and single_id not in requested_ids:
            requested_ids.append(single_id)
        profiles = (
            [_get_mailbox_or_404(mailbox_store, mailbox_id) for mailbox_id in requested_ids]
            if requested_ids
            else mailbox_store.list_mailboxes()
        )
        items = [
            {
                "mailbox": _to_jsonable(profile),
                "jobs": _to_jsonable(mailbox_store.list_sync_jobs(profile.id, limit=20)),
                "states": _to_jsonable(mailbox_store.list_sync_states(profile.id, method=_optional_text(payload.get("method")))),
            }
            for profile in profiles
        ]
        return jsonify({"items": items, "count": len(items)})

    @app.errorhandler(MailboxError)
    def handle_mailbox_error(error: MailboxError) -> Any:
        return jsonify({"error": {"code": error.code, "message": error.message}}), error.status_code

    @app.errorhandler(ValueError)
    def handle_value_error(error: ValueError) -> Any:
        return jsonify({"error": {"code": "invalid_request", "message": str(error)}}), 400

    @app.errorhandler(MailboxStoreError)
    def handle_store_error(error: MailboxStoreError) -> Any:
        message = str(error)
        status_code = 404 if "不存在" in message else 400
        code = "mailbox_not_found" if status_code == 404 else "mailbox_store_error"
        return jsonify({"error": {"code": code, "message": message}}), status_code

    @app.errorhandler(Exception)
    def handle_unexpected_error(error: Exception) -> Any:  # noqa: BLE001
        return jsonify({"error": {"code": "internal_error", "message": str(error)}}), 500

    return app


def _resolve_mailbox(store: MailboxStore, payload: dict[str, Any]) -> MailboxProfile:
    mailbox_id = payload.get("mailbox_id")
    if not isinstance(mailbox_id, int):
        raise MailboxError("请选择要操作的邮箱", code="missing_mailbox", status_code=400)
    return _get_mailbox_or_404(store, mailbox_id)


def _resolve_mailbox_by_email(store: MailboxStore, payload: dict[str, Any]) -> MailboxProfile:
    email = _require_text(payload, "email", "邮箱账号不能为空")
    mailbox = store.get_mailbox_by_email(email)
    if not mailbox:
        raise MailboxError("邮箱档案不存在", code="mailbox_not_found", status_code=404)
    return mailbox


def _get_mailbox_or_404(store: MailboxStore, mailbox_id: int) -> MailboxProfile:
    mailbox = store.get_mailbox(mailbox_id)
    if not mailbox:
        raise MailboxError("邮箱档案不存在", code="mailbox_not_found", status_code=404)
    return mailbox


def _profile_to_config(profile: MailboxProfile, *, method: str | None = None) -> MailboxConfig:
    return MailboxConfig(
        email=profile.email,
        client_id=profile.client_id,
        refresh_token=profile.refresh_token,
        proxy=profile.proxy,
        default_method=method or profile.preferred_method,
    )


def _profile_to_public_mailbox(profile: MailboxProfile) -> dict[str, Any]:
    return {
        "id": profile.id,
        "label": profile.label,
        "email": profile.email,
        "preferred_method": profile.preferred_method,
        "notes": profile.notes,
        "created_at": profile.created_at,
        "updated_at": profile.updated_at,
    }


def _build_runtime_config_from_payload(raw: dict[str, Any]) -> tuple[MailboxConfig, str]:
    method = _normalize_method(raw.get("preferred_method") or raw.get("method") or "graph_api")
    return (
        MailboxConfig(
            email=_require_text(raw, "email", "邮箱账号不能为空"),
            client_id=_require_text(raw, "client_id", "Client ID 不能为空"),
            refresh_token=_require_text(raw, "refresh_token", "Refresh Token 不能为空"),
            proxy=_optional_text(raw.get("proxy")),
            default_method=method,
        ),
        method,
    )


def _parse_positive_int(
    raw: Any,
    *,
    field_name: str,
    default: int,
    minimum: int = 1,
    maximum: int | None = None,
) -> int:
    if raw in (None, ""):
        return default
    try:
        value = int(raw)
    except (TypeError, ValueError) as exc:
        raise MailboxError(f"{field_name} 必须是整数", code=f"invalid_{field_name}") from exc
    if value < minimum:
        raise MailboxError(f"{field_name} 必须大于等于 {minimum}", code=f"invalid_{field_name}")
    if maximum is not None and value > maximum:
        raise MailboxError(f"{field_name} 不能大于 {maximum}", code=f"invalid_{field_name}")
    return value


def _parse_mailbox_ids(raw: Any, *, maximum: int) -> list[int]:
    if not isinstance(raw, list):
        raise MailboxError("mailbox_ids 必须是数组", code="invalid_mailbox_ids")
    if not raw:
        raise MailboxError("请至少选择一个邮箱", code="invalid_mailbox_ids")
    if len(raw) > maximum:
        raise MailboxError(f"一次最多处理 {maximum} 个邮箱", code="too_many_mailboxes")

    mailbox_ids: list[int] = []
    seen: set[int] = set()
    for item in raw:
        if not isinstance(item, int):
            raise MailboxError("mailbox_ids 中的每一项都必须是整数", code="invalid_mailbox_ids")
        if item in seen:
            continue
        seen.add(item)
        mailbox_ids.append(item)
    return mailbox_ids


def _parse_message_ids(raw: Any, *, maximum: int) -> list[str]:
    if not isinstance(raw, list):
        raise MailboxError("message_ids 必须是数组", code="invalid_message_ids")
    if not raw:
        raise MailboxError("请至少选择一封邮件", code="invalid_message_ids")
    if len(raw) > maximum:
        raise MailboxError(f"一次最多处理 {maximum} 封邮件", code="too_many_messages")

    message_ids: list[str] = []
    seen: set[str] = set()
    for item in raw:
        if not isinstance(item, str) or not item.strip():
            raise MailboxError("message_ids 中的每一项都必须是非空字符串", code="invalid_message_ids")
        normalized = item.strip()
        if normalized in seen:
            continue
        seen.add(normalized)
        message_ids.append(normalized)
    return message_ids


def _parse_optional_mailbox_ids(raw: Any) -> list[int]:
    if raw in (None, ""):
        return []
    return _parse_mailbox_ids(raw, maximum=500)


def _parse_optional_single_mailbox_id(raw: Any) -> int | None:
    if raw in (None, ""):
        return None
    return _parse_positive_int(raw, field_name="mailbox_id", default=1, minimum=1)


def _optional_text_list(raw: Any, *, field_name: str) -> list[str] | None:
    if raw is None:
        return None
    if not isinstance(raw, list):
        raise MailboxError(f"{field_name} 必须是字符串数组", code=f"invalid_{field_name}")
    result: list[str] = []
    for item in raw:
        if not isinstance(item, str) or not item.strip():
            raise MailboxError(f"{field_name} 中的每一项都必须是非空字符串", code=f"invalid_{field_name}")
        result.append(item.strip())
    return result


def _optional_object(raw: Any, *, field_name: str) -> dict[str, Any]:
    if raw in (None, ""):
        return {}
    if not isinstance(raw, dict):
        raise MailboxError(f"{field_name} 必须是对象", code=f"invalid_{field_name}")
    return raw


def _build_attachment_payload(raw: dict[str, Any]) -> dict[str, Any]:
    return {
        "name": _require_text(raw, "name", "缺少附件名称"),
        "content_type": _optional_text(raw.get("content_type")) or "application/octet-stream",
        "content_base64": _require_text(raw, "content_base64", "缺少附件内容"),
        "is_inline": _parse_bool(raw.get("is_inline"), field_name="is_inline", default=False),
        "content_id": _optional_text(raw.get("content_id")) or "",
    }


def _build_compose_payload(raw: dict[str, Any], *, require_to: bool) -> dict[str, Any]:
    payload = {
        "draft_message_id": _optional_text(raw.get("draft_message_id")) or _optional_text(raw.get("message_id")),
        "subject": _optional_text(raw.get("subject")) or "",
        "body_text": _optional_text(raw.get("body_text")) or _optional_text(raw.get("body")) or "",
        "body_html": _optional_text(raw.get("body_html")),
        "to_recipients": _optional_text_list(raw.get("to_recipients") or raw.get("to"), field_name="to_recipients") or [],
        "cc_recipients": _optional_text_list(raw.get("cc_recipients") or raw.get("cc"), field_name="cc_recipients") or [],
        "bcc_recipients": _optional_text_list(raw.get("bcc_recipients") or raw.get("bcc"), field_name="bcc_recipients") or [],
        "send_now": _parse_bool(raw.get("send_now"), field_name="send_now", default=True),
    }
    attachments = raw.get("attachments")
    if attachments is None:
        payload["attachments"] = []
    elif not isinstance(attachments, list):
        raise MailboxError("attachments 必须是数组", code="invalid_attachments")
    else:
        payload["attachments"] = []
        for item in attachments:
            if not isinstance(item, dict):
                raise MailboxError("attachments 中的每一项都必须是对象", code="invalid_attachments")
            payload["attachments"].append(_build_attachment_payload(item))

    if require_to and not payload["draft_message_id"] and payload["send_now"] and not payload["to_recipients"]:
        raise MailboxError("至少需要一个收件人", code="invalid_recipients")
    return payload


def _build_query(raw: dict[str, Any], *, default_method: str) -> MailboxQuery:
    method = _normalize_method(raw.get("method") or default_method)
    top = raw.get("top", 20)
    page = _parse_positive_int(raw.get("page"), field_name="page", default=1, minimum=1)
    page_size_raw = raw.get("page_size", top)
    page_size = _parse_positive_int(page_size_raw, field_name="page_size", default=20, minimum=1, maximum=100)
    unread_only = raw.get("unread_only", False)
    keyword = _optional_text(raw.get("keyword")) or ""
    folder = _optional_text(raw.get("folder")) or "INBOX"
    read_state = _normalize_read_state(raw.get("read_state"), unread_only=unread_only)
    has_attachments_only = _parse_bool(
        raw.get("has_attachments_only"),
        field_name="has_attachments_only",
        default=False,
    )
    flagged_only = _parse_bool(raw.get("flagged_only"), field_name="flagged_only", default=False)
    importance = _normalize_importance(raw.get("importance"))
    sort_order = _normalize_sort_order(raw.get("sort_order"))

    if not isinstance(top, int) or top < 1 or top > 50:
        raise MailboxError("拉取数量必须在 1 到 50 之间", code="invalid_top")
    if not isinstance(unread_only, bool):
        raise MailboxError("仅未读参数必须是布尔值", code="invalid_unread_only")
    return MailboxQuery(
        method=method,
        top=top,
        unread_only=unread_only,
        keyword=keyword,
        folder=folder,
        page=page,
        page_size=page_size,
        read_state=read_state,
        has_attachments_only=has_attachments_only,
        flagged_only=flagged_only,
        importance=importance,
        sort_order=sort_order,
    )


def _build_detail_request(raw: dict[str, Any], *, default_method: str) -> MailboxDetailRequest:
    method = _normalize_method(raw.get("method") or default_method)
    message_id = _require_text(raw, "message_id", "缺少邮件标识")
    folder = _optional_text(raw.get("folder")) or "INBOX"
    return MailboxDetailRequest(method=method, message_id=message_id, folder=folder)


def _build_read_state_request(raw: dict[str, Any], *, default_method: str) -> ReadStateUpdateRequest:
    method = _normalize_method(raw.get("method") or default_method)
    message_id = _require_text(raw, "message_id", "缺少邮件标识")
    is_read = raw.get("is_read")
    if not isinstance(is_read, bool):
        raise MailboxError("已读状态必须是布尔值", code="invalid_read_state")
    folder = _optional_text(raw.get("folder")) or "INBOX"
    return ReadStateUpdateRequest(method=method, message_id=message_id, is_read=is_read, folder=folder)


def _build_flag_state_request(raw: dict[str, Any], *, default_method: str) -> FlagStateUpdateRequest:
    method = _normalize_method(raw.get("method") or default_method)
    message_id = _require_text(raw, "message_id", "缺少邮件标识")
    is_flagged = raw.get("is_flagged")
    if not isinstance(is_flagged, bool):
        raise MailboxError("星标状态必须是布尔值", code="invalid_flag_state")
    folder = _optional_text(raw.get("folder")) or "INBOX"
    return FlagStateUpdateRequest(method=method, message_id=message_id, is_flagged=is_flagged, folder=folder)


def _build_move_request(raw: dict[str, Any], *, default_method: str) -> MessageMoveRequest:
    method = _normalize_method(raw.get("method") or default_method)
    message_id = _require_text(raw, "message_id", "缺少邮件标识")
    destination_folder = _require_text(raw, "destination_folder", "缺少目标文件夹")
    folder = _optional_text(raw.get("folder")) or "INBOX"
    return MessageMoveRequest(
        method=method,
        message_id=message_id,
        destination_folder=destination_folder,
        folder=folder,
    )


def _build_delete_request(raw: dict[str, Any], *, default_method: str) -> MessageDeleteRequest:
    method = _normalize_method(raw.get("method") or default_method)
    message_id = _require_text(raw, "message_id", "缺少邮件标识")
    folder = _optional_text(raw.get("folder")) or "INBOX"
    return MessageDeleteRequest(method=method, message_id=message_id, folder=folder)


def _extract_mailbox_payload(raw: dict[str, Any], *, partial: bool = False) -> dict[str, Any]:
    payload: dict[str, Any] = {}
    for key in ["label", "email", "client_id", "refresh_token", "proxy", "preferred_method", "notes"]:
        if key in raw:
            payload[key] = raw[key]

    if not partial:
        for key, message in {
            "label": "邮箱备注不能为空",
            "email": "邮箱账号不能为空",
            "client_id": "Client ID 不能为空",
            "refresh_token": "Refresh Token 不能为空",
        }.items():
            _require_text(raw, key, message)

    if "preferred_method" in payload:
        payload["preferred_method"] = _normalize_method(payload["preferred_method"])
    return payload


def _parse_import_mailboxes(
    raw_text: str,
    *,
    preferred_method: str,
    allow_missing_email: bool = False,
) -> list[dict[str, Any]]:
    content = raw_text.strip()
    if not content:
        raise MailboxError("批量导入文本不能为空", code="invalid_raw_text")

    if content[0] in "[{":
        return _parse_import_mailboxes_from_json(
            content,
            preferred_method=preferred_method,
            allow_missing_email=allow_missing_email,
        )
    if MAILBOX_IMPORT_DELIMITER in content:
        return _parse_import_mailboxes_from_delimited_lines(
            content,
            preferred_method=preferred_method,
            allow_missing_email=allow_missing_email,
        )
    return _parse_import_mailboxes_from_tabular_text(
        content,
        preferred_method=preferred_method,
        allow_missing_email=allow_missing_email,
    )


def _parse_import_mailboxes_from_json(
    raw_text: str,
    *,
    preferred_method: str,
    allow_missing_email: bool = False,
) -> list[dict[str, Any]]:
    try:
        parsed = json.loads(raw_text)
    except json.JSONDecodeError as error:
        raise MailboxError(
            f"JSON 解析失败：第 {error.lineno} 行第 {error.colno} 列附近格式不正确",
            code="invalid_import_line",
        ) from error

    if isinstance(parsed, dict):
        records: Any = parsed.get("mailboxes")
        if records is None:
            records = parsed.get("items")
        if records is None:
            records = [parsed]
    else:
        records = parsed

    if not isinstance(records, list):
        raise MailboxError("JSON 导入内容必须是对象数组", code="invalid_import_line")

    payloads = [
        _normalize_import_mailbox_record(
            record,
            preferred_method=preferred_method,
            row_number=index,
            allow_missing_email=allow_missing_email,
        )
        for index, record in enumerate(records, start=1)
        if isinstance(record, dict)
    ]
    if len(payloads) != len(records):
        raise MailboxError("JSON 导入的每一项都必须是对象", code="invalid_import_line")
    if not payloads:
        raise MailboxError("批量导入文本不能为空", code="invalid_raw_text")
    return payloads


def _parse_import_mailboxes_from_delimited_lines(
    raw_text: str,
    *,
    preferred_method: str,
    allow_missing_email: bool = False,
) -> list[dict[str, Any]]:
    payloads: list[dict[str, Any]] = []

    for line_number, raw_line in enumerate(raw_text.splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue

        payloads.append(
            _parse_delimited_import_record(
                line,
                preferred_method=preferred_method,
                row_number=line_number,
                allow_missing_email=allow_missing_email,
            )
        )

    if not payloads:
        raise MailboxError("批量导入文本不能为空", code="invalid_raw_text")
    return payloads


def _parse_delimited_import_record(
    line: str,
    *,
    preferred_method: str,
    row_number: int,
    allow_missing_email: bool = False,
) -> dict[str, Any]:
    parts = [part.strip() for part in line.split(MAILBOX_IMPORT_DELIMITER) if part.strip()]
    if len(parts) == 2:
        if allow_missing_email and not any(_looks_like_email(part) for part in parts):
            client_id, refresh_token = _resolve_client_and_refresh_values(parts)
            return _normalize_import_mailbox_record(
                {
                    "client_id": client_id,
                    "refresh_token": refresh_token,
                },
                preferred_method=preferred_method,
                row_number=row_number,
                allow_missing_email=True,
            )
        raise MailboxError(
            f"第 {row_number} 行格式错误，若只传两段内容，必须是 client_id 和 refresh_token",
            code="invalid_import_line",
        )
    if len(parts) < 3 or len(parts) > 4:
        raise MailboxError(
            (
                f"第 {row_number} 行格式错误，支持 email----client_id----refresh_token、"
                "email----附加字段----client_id----refresh_token，"
                "也支持 key=value 写法或包含邮箱时的灵活顺序"
            ),
            code="invalid_import_line",
        )

    keyed_record: dict[str, str] = {}
    anonymous_parts: list[str] = []
    for part in parts:
        parsed_segment = _parse_import_keyed_segment(part)
        if parsed_segment:
            field_name, value = parsed_segment
            keyed_record[field_name] = value
        else:
            anonymous_parts.append(part)

    if not keyed_record and len(parts) == 3 and _looks_like_email(parts[0]):
        return _normalize_import_mailbox_record(
            {
                "email": parts[0],
                "client_id": parts[1],
                "refresh_token": parts[2],
            },
            preferred_method=preferred_method,
            row_number=row_number,
            allow_missing_email=allow_missing_email,
        )

    if not keyed_record and len(parts) == 4 and _looks_like_email(parts[0]):
        return _normalize_import_mailbox_record(
            {
                "email": parts[0],
                "client_id": parts[2],
                "refresh_token": parts[3],
            },
            preferred_method=preferred_method,
            row_number=row_number,
            allow_missing_email=allow_missing_email,
        )

    email_candidates = [value for value in anonymous_parts if _looks_like_email(value)]
    if "email" not in keyed_record and len(email_candidates) == 1:
        keyed_record["email"] = email_candidates[0]
        anonymous_parts.remove(email_candidates[0])

    method_candidates = [value for value in anonymous_parts if _looks_like_method(value)]
    if "preferred_method" not in keyed_record and len(method_candidates) == 1:
        keyed_record["preferred_method"] = method_candidates[0]
        anonymous_parts.remove(method_candidates[0])

    missing_required_fields = [
        field_name
        for field_name in ("email", "client_id", "refresh_token")
        if field_name not in keyed_record
    ]

    if not keyed_record.get("email") and not allow_missing_email:
        raise MailboxError(
            f"第 {row_number} 行缺少邮箱账号，导入至少需要 email、client_id、refresh_token",
            code="invalid_import_line",
        )

    if missing_required_fields == ["client_id", "refresh_token"] and len(anonymous_parts) == 2:
        client_id, refresh_token = _resolve_client_and_refresh_values(anonymous_parts)
        keyed_record["client_id"] = client_id
        keyed_record["refresh_token"] = refresh_token
    elif len(missing_required_fields) == 1 and len(anonymous_parts) == 1:
        keyed_record[missing_required_fields[0]] = anonymous_parts[0]
    elif len(missing_required_fields) == 2 and len(anonymous_parts) == 3:
        client_id, refresh_token = _resolve_client_and_refresh_values(anonymous_parts[-2:])
        keyed_record["client_id"] = client_id
        keyed_record["refresh_token"] = refresh_token
    elif anonymous_parts and not keyed_record.get("client_id") and not keyed_record.get("refresh_token"):
        client_id, refresh_token = _resolve_client_and_refresh_values(anonymous_parts[-2:])
        keyed_record["client_id"] = client_id
        keyed_record["refresh_token"] = refresh_token

    return _normalize_import_mailbox_record(
        keyed_record,
        preferred_method=preferred_method,
        row_number=row_number,
        allow_missing_email=allow_missing_email,
    )


def _parse_import_keyed_segment(segment: str) -> tuple[str, str] | None:
    for delimiter in ("=", ":", "："):
        if delimiter not in segment:
            continue
        raw_key, raw_value = segment.split(delimiter, 1)
        field_name = _resolve_import_field_name(raw_key)
        if field_name and raw_value.strip():
            return field_name, raw_value.strip()
    return None


def _resolve_client_and_refresh_values(values: list[str]) -> tuple[str, str]:
    if len(values) != 2:
        raise MailboxError("导入记录缺少 Client ID 或 Refresh Token", code="invalid_import_line")

    first, second = values
    if _looks_like_client_id(first) and not _looks_like_client_id(second):
        return first, second
    if _looks_like_client_id(second) and not _looks_like_client_id(first):
        return second, first
    if _looks_like_refresh_token(first) and not _looks_like_refresh_token(second):
        return second, first
    if _looks_like_refresh_token(second) and not _looks_like_refresh_token(first):
        return first, second
    return first, second


def _parse_import_mailboxes_from_tabular_text(
    raw_text: str,
    *,
    preferred_method: str,
    allow_missing_email: bool = False,
) -> list[dict[str, Any]]:
    reader = csv.reader(StringIO(raw_text), delimiter=_detect_tabular_import_delimiter(raw_text))
    rows = [[cell.strip() for cell in row] for row in reader if any(cell.strip() for cell in row)]
    if not rows:
        raise MailboxError("批量导入文本不能为空", code="invalid_raw_text")

    headers = rows[0] if _looks_like_import_header(rows[0]) else None
    start_index = 2 if headers else 1
    data_rows = rows[1:] if headers else rows

    payloads: list[dict[str, Any]] = []
    for offset, row in enumerate(data_rows, start=start_index):
        if headers:
            if len(row) != len(headers):
                raise MailboxError(f"第 {offset} 行列数与表头不一致", code="invalid_import_line")
            raw_record = {
                field_name: value
                for header, value in zip(headers, row)
                if (field_name := _resolve_import_field_name(header))
            }
        else:
            if len(row) == 3 and all(row):
                if allow_missing_email and not any(_looks_like_email(value) for value in row):
                    client_id, refresh_token = _resolve_client_and_refresh_values([row[0], row[1]])
                    raw_record = {
                        "client_id": client_id,
                        "refresh_token": refresh_token,
                        "preferred_method": row[2] if _looks_like_method(row[2]) else None,
                    }
                else:
                    raw_record = {
                        "email": row[0],
                        "client_id": row[1],
                        "refresh_token": row[2],
                    }
            elif len(row) == 2 and all(row) and allow_missing_email and not any(_looks_like_email(value) for value in row):
                client_id, refresh_token = _resolve_client_and_refresh_values(row)
                raw_record = {
                    "client_id": client_id,
                    "refresh_token": refresh_token,
                }
            elif len(row) == 4 and all(row):
                raw_record = {
                    "email": row[0],
                    "client_id": row[2],
                    "refresh_token": row[3],
                }
            else:
                raise MailboxError(
                    (
                        f"第 {offset} 行格式错误，CSV/TSV 无表头时请使用 "
                        "邮箱,Client ID,Refresh Token 或 "
                        "邮箱,附加字段,Client ID,Refresh Token"
                    ),
                    code="invalid_import_line",
                )

        payloads.append(
            _normalize_import_mailbox_record(
                raw_record,
                preferred_method=preferred_method,
                row_number=offset,
                allow_missing_email=allow_missing_email,
            )
        )

    if not payloads:
        raise MailboxError("批量导入文本不能为空", code="invalid_raw_text")
    return payloads


def _detect_tabular_import_delimiter(raw_text: str) -> str:
    sample = "\n".join(line for line in raw_text.splitlines() if line.strip()) or raw_text
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=MAILBOX_IMPORT_TABULAR_DELIMITERS)
        delimiter = getattr(dialect, "delimiter", "")
        if delimiter in MAILBOX_IMPORT_TABULAR_DELIMITERS:
            return delimiter
    except csv.Error:
        pass

    delimiter = max(
        MAILBOX_IMPORT_TABULAR_DELIMITERS,
        key=lambda current: sum(current in line for line in raw_text.splitlines()),
    )
    if delimiter and delimiter in raw_text:
        return delimiter
    raise MailboxError(
        "批量导入格式不正确，支持 JSON、CSV/TSV 或 ---- 分隔文本",
        code="invalid_import_line",
    )


def _looks_like_import_header(row: list[str]) -> bool:
    resolved_fields = {_resolve_import_field_name(value) for value in row}
    return {"client_id", "refresh_token"}.issubset(resolved_fields)


def _hydrate_import_mailboxes_missing_email(
    manager: MailboxManager,
    payloads: list[dict[str, Any]],
    *,
    preferred_method: str,
) -> list[dict[str, Any]]:
    hydrated_payloads: list[dict[str, Any]] = []
    for index, payload in enumerate(payloads, start=1):
        email = _optional_text(payload.get("email"))
        if email:
            hydrated_payloads.append(payload)
            continue

        client_id = _optional_text(payload.get("client_id"))
        refresh_token = _optional_text(payload.get("refresh_token"))
        proxy = _optional_text(payload.get("proxy"))
        if not client_id or not refresh_token:
            raise MailboxError(
                f"第 {index} 条记录缺少 Client ID 或 Refresh Token，无法自动补全邮箱",
                code="invalid_import_line",
            )

        resolved_email = manager.resolve_mailbox_email(
            client_id=client_id,
            refresh_token=refresh_token,
            proxy=proxy,
        )
        hydrated_payloads.append(
            {
                **payload,
                "email": resolved_email,
                "preferred_method": payload.get("preferred_method") or preferred_method,
            }
        )
    return hydrated_payloads


def _normalize_import_mailbox_record(
    raw: dict[str, Any],
    *,
    preferred_method: str,
    row_number: int,
    allow_missing_email: bool = False,
) -> dict[str, Any]:
    if not isinstance(raw, dict):
        raise MailboxError(f"第 {row_number} 条记录必须是对象", code="invalid_import_line")

    normalized_raw = _canonicalize_import_record(raw)
    email = _optional_text(normalized_raw.get("email"))
    client_id = _optional_text(normalized_raw.get("client_id"))
    refresh_token = _optional_text(normalized_raw.get("refresh_token"))
    method_value = normalized_raw.get("preferred_method")

    if not email and not allow_missing_email:
        raise MailboxError(f"第 {row_number} 条记录缺少邮箱账号", code="invalid_import_line")
    if not client_id:
        raise MailboxError(f"第 {row_number} 条记录缺少 Client ID", code="invalid_import_line")
    if not refresh_token:
        raise MailboxError(f"第 {row_number} 条记录缺少 Refresh Token", code="invalid_import_line")

    return {
        "email": email,
        "client_id": client_id,
        "refresh_token": refresh_token,
        "preferred_method": _normalize_method(method_value or preferred_method),
    }


def _canonicalize_import_record(raw: dict[str, Any]) -> dict[str, Any]:
    normalized_raw: dict[str, Any] = {}
    for key, value in raw.items():
        if isinstance(key, str):
            field_name = _resolve_import_field_name(key)
            normalized_raw[field_name or key] = value
        else:
            normalized_raw[key] = value
    return normalized_raw


def _resolve_import_field_name(value: str) -> str | None:
    normalized = (
        value.strip()
        .casefold()
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .replace(".", "")
        .replace("/", "")
        .replace("：", "")
        .replace(":", "")
    )
    return {
        "email": "email",
        "mail": "email",
        "mailbox": "email",
        "account": "email",
        "邮箱": "email",
        "邮箱账号": "email",
        "邮箱地址": "email",
        "clientid": "client_id",
        "clinetid": "client_id",
        "client": "client_id",
        "cid": "client_id",
        "应用id": "client_id",
        "refreshtoken": "refresh_token",
        "refresh": "refresh_token",
        "token": "refresh_token",
        "rt": "refresh_token",
        "preferredmethod": "preferred_method",
        "method": "preferred_method",
        "默认方法": "preferred_method",
    }.get(normalized)


def _looks_like_email(value: str) -> bool:
    return bool(EMAIL_PATTERN.match(value.strip()))


def _looks_like_client_id(value: str) -> bool:
    cleaned = value.strip()
    if CLIENT_ID_PATTERN.match(cleaned):
        return True
    return bool(cleaned) and len(cleaned) <= 80 and all(char.isalnum() or char in "-_" for char in cleaned)


def _looks_like_refresh_token(value: str) -> bool:
    cleaned = value.strip()
    if not cleaned:
        return False
    if cleaned.startswith("M."):
        return True
    punctuation_count = sum(char in ".*!_-$=" for char in cleaned)
    return len(cleaned) >= 24 and punctuation_count >= 1


def _looks_like_method(value: str) -> bool:
    cleaned = value.strip()
    return cleaned in {"graph_api", "graph", "imap_new", "imap_old"}


def _normalize_method(value: Any) -> str:
    if not isinstance(value, str) or not value.strip():
        raise MailboxError("请选择邮件接入方式", code="invalid_method")
    method = value.strip()
    if method == "graph":
        method = "graph_api"
    if method not in {"graph_api", "imap_new", "imap_old"}:
        raise MailboxError("不支持的邮件接入方式", code="invalid_method")
    return method


def _normalize_read_state(value: Any, *, unread_only: bool = False) -> str:
    if value in (None, ""):
        return "unread" if unread_only else "all"
    if not isinstance(value, str):
        raise MailboxError("read_state 参数类型错误", code="invalid_read_state")
    normalized = value.strip().casefold()
    if normalized not in {"all", "read", "unread"}:
        raise MailboxError("read_state 仅支持 all、read、unread", code="invalid_read_state")
    return normalized


def _normalize_importance(value: Any) -> str:
    if value in (None, ""):
        return "all"
    if not isinstance(value, str):
        raise MailboxError("importance 参数类型错误", code="invalid_importance")
    normalized = value.strip().casefold()
    if normalized not in {"all", "high", "normal", "low"}:
        raise MailboxError("importance 仅支持 all、high、normal、low", code="invalid_importance")
    return normalized


def _normalize_sort_order(value: Any) -> str:
    if value in (None, ""):
        return "desc"
    if not isinstance(value, str):
        raise MailboxError("sort_order 参数类型错误", code="invalid_sort_order")
    normalized = value.strip().casefold()
    if normalized not in {"asc", "desc"}:
        raise MailboxError("sort_order 仅支持 asc、desc", code="invalid_sort_order")
    return normalized


def _normalize_message_action(value: Any) -> str:
    if not isinstance(value, str) or not value.strip():
        raise MailboxError("缺少批量操作类型", code="invalid_message_action")
    normalized = value.strip().casefold()
    if normalized not in {"mark_read", "mark_unread", "flag", "unflag", "delete", "archive", "move"}:
        raise MailboxError("不支持的批量操作类型", code="invalid_message_action")
    return normalized


def _parse_bool(raw: Any, *, field_name: str, default: bool) -> bool:
    if raw is None:
        return default
    if isinstance(raw, bool):
        return raw
    raise MailboxError(f"{field_name} 必须是布尔值", code=f"invalid_{field_name}")


def _require_text(raw: dict[str, Any], key: str, error_message: str) -> str:
    value = _optional_text(raw.get(key))
    if not value:
        raise MailboxError(error_message, code=f"invalid_{key}")
    return value


def _optional_text(value: Any) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str):
        raise MailboxError("请求参数类型错误", code="invalid_payload")
    cleaned = value.strip()
    return cleaned or None


def _message_value(message: Any, key: str, default: Any = None) -> Any:
    if isinstance(message, dict):
        return message.get(key, default)
    return getattr(message, key, default)


def _rule_matches_message(message: Any, conditions: dict[str, Any]) -> bool:
    sender_contains = _optional_text(conditions.get("sender_contains")) if isinstance(conditions.get("sender_contains"), str) else None
    subject_contains = _optional_text(conditions.get("subject_contains")) if isinstance(conditions.get("subject_contains"), str) else None
    folder = _optional_text(conditions.get("folder")) if isinstance(conditions.get("folder"), str) else None
    importance = _optional_text(conditions.get("importance")) if isinstance(conditions.get("importance"), str) else None
    keyword = _optional_text(conditions.get("keyword")) if isinstance(conditions.get("keyword"), str) else None
    required_tag = _optional_text(conditions.get("tag")) if isinstance(conditions.get("tag"), str) else None

    sender_text = f"{_message_value(message, 'sender', '')} {_message_value(message, 'sender_name', '')}".casefold()
    subject_text = str(_message_value(message, "subject", "") or "").casefold()
    preview_text = str(_message_value(message, "preview", "") or "").casefold()

    if sender_contains and sender_contains.casefold() not in sender_text:
        return False
    if subject_contains and subject_contains.casefold() not in subject_text:
        return False
    if folder and folder.casefold() != str(_message_value(message, "folder", "") or "").casefold():
        return False
    if importance and importance.casefold() != str(_message_value(message, "importance", "") or "").casefold():
        return False
    if keyword and keyword.casefold() not in f"{subject_text} {sender_text} {preview_text}":
        return False
    if conditions.get("has_attachments") is True and not bool(_message_value(message, "has_attachments", False)):
        return False
    if conditions.get("is_unread") is True and bool(_message_value(message, "is_read", False)):
        return False
    if conditions.get("is_flagged") is True and not bool(_message_value(message, "is_flagged", False)):
        return False
    if required_tag:
        meta = _message_value(message, "meta")
        tags = list(getattr(meta, "tags", [])) if meta else []
        if required_tag not in tags:
            return False
    return True


def _apply_rule_actions_to_message(
    manager: MailboxManager,
    store: MailboxStore,
    profile: MailboxProfile,
    message: Any,
    actions: dict[str, Any],
    *,
    method: str | None = None,
) -> None:
    resolved_method = method or str(_message_value(message, "method", profile.preferred_method))
    message_id = str(_message_value(message, "message_id", "") or "")
    folder = str(_message_value(message, "folder", "") or "INBOX")
    config = _profile_to_config(profile, method=resolved_method)

    if actions.get("mark_read") is True:
        manager.update_read_state(
            config,
            ReadStateUpdateRequest(method=resolved_method, message_id=message_id, is_read=True, folder=folder),
        )
        store.update_cached_message_state(profile.id, resolved_method, message_id, is_read=True)

    if "flag" in actions:
        is_flagged = bool(actions.get("flag"))
        manager.update_flag_state(
            config,
            FlagStateUpdateRequest(method=resolved_method, message_id=message_id, is_flagged=is_flagged, folder=folder),
        )
        store.update_cached_message_state(profile.id, resolved_method, message_id, is_flagged=is_flagged)

    move_to_folder = _optional_text(actions.get("move_to_folder")) if isinstance(actions.get("move_to_folder"), str) else None
    if move_to_folder:
        manager.move_message(
            config,
            MessageMoveRequest(
                method=resolved_method,
                message_id=message_id,
                destination_folder=move_to_folder,
                folder=folder,
            ),
        )
        store.update_cached_message_state(profile.id, resolved_method, message_id, folder_id=move_to_folder)

    existing_meta = store.get_message_meta(profile.id, resolved_method, message_id)
    next_tags = list(existing_meta.tags) if existing_meta else []
    tags_value = actions.get("tags")
    if isinstance(tags_value, str) and tags_value.strip():
        if tags_value.strip() not in next_tags:
            next_tags.append(tags_value.strip())
    elif isinstance(tags_value, list):
        for item in tags_value:
            if isinstance(item, str) and item.strip() and item.strip() not in next_tags:
                next_tags.append(item.strip())

    notes_append = _optional_text(actions.get("notes_append")) if isinstance(actions.get("notes_append"), str) else None
    next_notes = (existing_meta.notes if existing_meta else "") if existing_meta else ""
    if notes_append:
        next_notes = f"{next_notes}\n{notes_append}".strip() if next_notes else notes_append

    follow_up = _optional_text(actions.get("follow_up")) if isinstance(actions.get("follow_up"), str) else None
    snoozed_until = _optional_text(actions.get("snoozed_until")) if isinstance(actions.get("snoozed_until"), str) else None
    status = _optional_text(actions.get("status")) if isinstance(actions.get("status"), str) else None

    if next_tags or notes_append or follow_up is not None or snoozed_until is not None or status is not None:
        store.update_message_meta(
            profile.id,
            resolved_method,
            message_id,
            tags=next_tags if next_tags else None,
            follow_up=follow_up,
            notes=next_notes if (notes_append or existing_meta) else None,
            snoozed_until=snoozed_until,
            status=status,
        )


def _apply_rules_to_messages(
    manager: MailboxManager,
    store: MailboxStore,
    profile: MailboxProfile,
    rules: list[Any],
    messages: list[Any],
    *,
    method: str | None = None,
) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    for rule in rules:
        if not rule:
            continue
        matched = 0
        applied = 0
        failed = 0
        for message in messages:
            if not _rule_matches_message(message, getattr(rule, "conditions", {})):
                continue
            matched += 1
            try:
                _apply_rule_actions_to_message(
                    manager,
                    store,
                    profile,
                    message,
                    getattr(rule, "actions", {}),
                    method=method,
                )
                applied += 1
            except Exception:  # noqa: BLE001
                failed += 1
        store.record_audit_log(
            mailbox_id=profile.id,
            actor="admin",
            action="rule_applied",
            target_type="rule",
            target_id=str(getattr(rule, "id", "")),
            status="success" if failed == 0 else "partial",
            details={"matched": matched, "applied": applied, "failed": failed},
        )
        results.append(
            {
                "rule_id": getattr(rule, "id", None),
                "name": getattr(rule, "name", ""),
                "matched": matched,
                "applied": applied,
                "failed": failed,
            }
        )
    return results


def _extract_access_key_from_request() -> str | None:
    header_value = _optional_text(request.headers.get("X-InboxOps-Key"))
    if header_value:
        return header_value

    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return None
    return _optional_text(payload.get("api_key") or payload.get("key"))


def _probe_mailbox_connection(
    manager: MailboxManager,
    *,
    config: MailboxConfig,
    method: str,
) -> str:
    result = manager.list_messages(
        config,
        MailboxQuery(method=method, top=1, unread_only=False, keyword=""),
    )
    messages = result.messages if hasattr(result, "messages") else result
    latest_message = messages[0] if messages else None
    latest_subject = getattr(latest_message, "subject", "") if latest_message else ""
    if not latest_subject and isinstance(latest_message, dict):
        latest_subject = str(latest_message.get("subject", ""))
    return "连接成功，收件箱暂无邮件" if not latest_subject else f"连接成功，最近邮件：{latest_subject}"


def _label_for_method(method: str) -> str:
    return {
        "graph_api": "Graph API",
        "imap_new": "新版 IMAP",
        "imap_old": "旧版 IMAP",
    }.get(method, method)


def _to_jsonable(value: Any) -> Any:
    if is_dataclass(value):
        return {key: _to_jsonable(item) for key, item in asdict(value).items()}
    if isinstance(value, list):
        return [_to_jsonable(item) for item in value]
    if isinstance(value, tuple):
        return [_to_jsonable(item) for item in value]
    if isinstance(value, dict):
        return {key: _to_jsonable(item) for key, item in value.items()}
    return value


app = create_app()


if __name__ == "__main__":
    app.run(debug=True)

