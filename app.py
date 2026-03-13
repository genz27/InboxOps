from __future__ import annotations

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
    MailboxConfig,
    MailboxDetailRequest,
    MailboxError,
    MailboxManager,
    MailboxQuery,
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
        return jsonify({"mailbox": _to_jsonable(profile), "method": query.method, "messages": _to_jsonable(messages), "count": count})

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
        return jsonify(
            {
                "mailbox": _profile_to_public_mailbox(profile),
                "method": query.method,
                "messages": _to_jsonable(messages),
                "count": count,
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
        return jsonify({"mailbox": _to_jsonable(profile), "message": _to_jsonable(message)})

    @app.post("/api/key/mailbox/message")
    @api_key_required
    def mailbox_message_by_key() -> Any:
        payload = request.get_json(silent=True) or {}
        profile = _resolve_mailbox_by_email(mailbox_store, payload)
        detail_request = _build_detail_request(payload, default_method=profile.preferred_method)
        config = _profile_to_config(profile, method=detail_request.method)
        message = mailbox_manager.get_message_detail(config, detail_request)
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
        result = mailbox_manager.update_read_state(config, state_request)
        if hasattr(result, "body_text"):
            message = result
        else:
            message = mailbox_manager.get_message_detail(
                config,
                MailboxDetailRequest(method=state_request.method, message_id=state_request.message_id),
            )
        return jsonify({"mailbox": _to_jsonable(profile), "message": _to_jsonable(message)})

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


def _build_query(raw: dict[str, Any], *, default_method: str) -> MailboxQuery:
    method = _normalize_method(raw.get("method") or default_method)
    top = raw.get("top", 10)
    unread_only = raw.get("unread_only", False)
    keyword = _optional_text(raw.get("keyword")) or ""

    if not isinstance(top, int) or top < 1 or top > 50:
        raise MailboxError("拉取数量必须在 1 到 50 之间", code="invalid_top")
    if not isinstance(unread_only, bool):
        raise MailboxError("仅未读参数必须是布尔值", code="invalid_unread_only")
    return MailboxQuery(method=method, top=top, unread_only=unread_only, keyword=keyword)


def _build_detail_request(raw: dict[str, Any], *, default_method: str) -> MailboxDetailRequest:
    method = _normalize_method(raw.get("method") or default_method)
    message_id = _require_text(raw, "message_id", "缺少邮件标识")
    return MailboxDetailRequest(method=method, message_id=message_id)


def _build_read_state_request(raw: dict[str, Any], *, default_method: str) -> ReadStateUpdateRequest:
    method = _normalize_method(raw.get("method") or default_method)
    message_id = _require_text(raw, "message_id", "缺少邮件标识")
    is_read = raw.get("is_read")
    if not isinstance(is_read, bool):
        raise MailboxError("已读状态必须是布尔值", code="invalid_read_state")
    return ReadStateUpdateRequest(method=method, message_id=message_id, is_read=is_read)


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

