from __future__ import annotations

import base64

from app import create_app
from services.outlook_manager import MailboxError, MessageDetail, MessageListResult, MessageSummary
from services.storage import MailboxStore


class RecordingManager:
    def __init__(
        self,
        *,
        fail_emails: set[str] | None = None,
        resolved_emails: dict[tuple[str, str], str] | None = None,
        resolve_email_error: str | None = None,
    ) -> None:
        self.fail_emails = fail_emails or set()
        self.resolved_emails = resolved_emails or {}
        self.resolve_email_error = resolve_email_error
        self.calls: list[tuple[object, object]] = []
        self.detail_calls: list[tuple[object, object]] = []
        self.folder_calls: list[tuple[object, object]] = []
        self.read_calls: list[tuple[object, object]] = []
        self.flag_calls: list[tuple[object, object]] = []
        self.move_calls: list[tuple[object, object]] = []
        self.delete_calls: list[tuple[object, object]] = []
        self.draft_calls: list[tuple[object, object]] = []
        self.send_calls: list[tuple[object, object]] = []
        self.reply_calls: list[tuple[object, object, object, bool]] = []
        self.forward_calls: list[tuple[object, object, object]] = []
        self.upload_attachment_calls: list[tuple[object, object, object]] = []
        self.download_attachment_calls: list[tuple[object, object, object]] = []
        self.folder_create_calls: list[tuple[object, object]] = []
        self.folder_rename_calls: list[tuple[object, object]] = []
        self.folder_delete_calls: list[tuple[object, object]] = []

    def list_messages(self, config: object, request: object) -> object:
        self.calls.append((config, request))
        if getattr(config, "email", "") in self.fail_emails:
            raise MailboxError("连接失败")
        method = getattr(request, "method", "graph_api")
        return MessageListResult(
            method=method,
            total=1,
            returned=1,
            folder=getattr(request, "folder", "INBOX"),
            page=getattr(request, "page", 1),
            page_size=getattr(request, "page_size", 20),
            total_pages=1,
            has_prev=False,
            has_next=False,
            messages=[
                MessageSummary(
                    method=method,
                    message_id="msg-001",
                    subject="Welcome mail",
                    sender="sender@example.com",
                    sender_name="Sender",
                    received_at="2026-03-13T00:00:00Z",
                    is_read=False,
                    has_attachments=True,
                    preview="Preview",
                    source="Graph API",
                    folder=getattr(request, "folder", "INBOX"),
                    is_flagged=True,
                    importance="high",
                    conversation_id="conv-1",
                )
            ],
        )

    def list_folders(self, config: object, method: str) -> list[dict[str, object]]:
        self.folder_calls.append((config, method))
        return [
            {
                "id": "INBOX",
                "name": "inbox",
                "display_name": "收件箱",
                "total": 12,
                "unread": 3,
                "is_default": True,
                "kind": "inbox",
            },
            {
                "id": "archive",
                "name": "archive",
                "display_name": "归档",
                "total": 5,
                "unread": 0,
                "is_default": False,
                "kind": "archive",
            },
        ]

    def get_message_detail(self, config: object, request: object) -> object:
        self.detail_calls.append((config, request))
        if getattr(config, "email", "") in self.fail_emails:
            raise MailboxError("连接失败")
        return MessageDetail(
            method=getattr(request, "method", "graph_api"),
            message_id=request.message_id,
            subject="Detail subject",
            sender="sender@example.com",
            sender_name="Sender",
            received_at="2026-03-13T00:00:00Z",
            is_read=False,
            has_attachments=True,
            preview="Preview",
            source="Graph API",
            folder=getattr(request, "folder", "INBOX"),
            internet_message_id="<demo>",
            body_text="Hello detail",
            body_html="<p>Hello <strong>detail</strong></p>",
            to_recipients=["demo@example.com"],
            cc_recipients=[],
            bcc_recipients=[],
            headers={"Message-ID": "<demo>"},
            attachments=[],
            is_flagged=True,
            importance="high",
            conversation_id="conv-1",
        )

    def update_flag_state(self, config: object, request: object) -> object:
        self.flag_calls.append((config, request))
        return {
            "method": getattr(request, "method", "graph_api"),
            "message_id": request.message_id,
            "is_flagged": bool(request.is_flagged),
            "status": "updated",
        }

    def update_read_state(self, config: object, request: object) -> object:
        self.read_calls.append((config, request))
        return {
            "method": getattr(request, "method", "graph_api"),
            "message_id": request.message_id,
            "is_read": bool(request.is_read),
            "status": "updated",
        }

    def move_message(self, config: object, request: object) -> object:
        self.move_calls.append((config, request))
        return {
            "method": getattr(request, "method", "graph_api"),
            "message_id": request.message_id,
            "source_folder": getattr(request, "folder", "INBOX"),
            "destination_folder": request.destination_folder,
            "status": "moved",
        }

    def delete_message(self, config: object, request: object) -> object:
        self.delete_calls.append((config, request))
        return {
            "method": getattr(request, "method", "graph_api"),
            "message_id": request.message_id,
            "folder": getattr(request, "folder", "INBOX"),
            "status": "deleted",
        }

    def save_draft(self, config: object, payload: object) -> object:
        self.draft_calls.append((config, payload))
        attachments = [
            {
                "id": item.get("id") or f"draft-att-{index}",
                "name": item["name"],
                "content_type": item.get("content_type", "application/octet-stream"),
                "size": len(base64.b64decode(item["content_base64"])),
                "is_inline": bool(item.get("is_inline", False)),
                "content_base64": item["content_base64"],
                "content_id": item.get("content_id", ""),
            }
            for index, item in enumerate(payload.get("attachments", []), start=1)
        ]
        return {
            "method": getattr(config, "default_method", "graph_api"),
            "message_id": "draft-001",
            "subject": payload.get("subject", "Draft subject"),
            "sender": getattr(config, "email", ""),
            "sender_name": "Draft Sender",
            "received_at": "2026-03-13T00:00:00Z",
            "is_read": True,
            "has_attachments": bool(payload.get("attachments")),
            "preview": payload.get("body_text", "Draft body"),
            "source": "Graph API",
            "folder": "drafts",
            "internet_message_id": "<draft-001>",
            "is_flagged": False,
            "importance": "normal",
            "conversation_id": "conv-draft",
            "body_text": payload.get("body_text", "Draft body"),
            "body_html": payload.get("body_html"),
            "to": payload.get("to_recipients", []),
            "cc": payload.get("cc_recipients", []),
            "bcc": payload.get("bcc_recipients", []),
            "headers": {"Message-ID": "<draft-001>"},
            "attachments": attachments,
            "status": "draft_saved",
        }

    def send_message(self, config: object, payload: object) -> object:
        self.send_calls.append((config, payload))
        attachments = [
            {
                "id": item.get("id") or f"sent-att-{index}",
                "name": item["name"],
                "content_type": item.get("content_type", "application/octet-stream"),
                "size": len(base64.b64decode(item["content_base64"])),
                "is_inline": bool(item.get("is_inline", False)),
                "content_base64": item["content_base64"],
                "content_id": item.get("content_id", ""),
            }
            for index, item in enumerate(payload.get("attachments", []), start=1)
        ]
        return {
            "method": getattr(config, "default_method", "graph_api"),
            "message_id": "sent-001",
            "subject": payload.get("subject", "Sent subject"),
            "sender": getattr(config, "email", ""),
            "sender_name": "Sender",
            "received_at": "2026-03-13T00:00:00Z",
            "is_read": True,
            "has_attachments": bool(payload.get("attachments")),
            "preview": payload.get("body_text", "Sent body"),
            "source": "Graph API",
            "folder": "sentitems",
            "internet_message_id": "<sent-001>",
            "is_flagged": False,
            "importance": "normal",
            "conversation_id": "conv-sent",
            "body_text": payload.get("body_text", "Sent body"),
            "body_html": payload.get("body_html"),
            "to": payload.get("to_recipients", []),
            "cc": payload.get("cc_recipients", []),
            "bcc": payload.get("bcc_recipients", []),
            "headers": {"Message-ID": "<sent-001>"},
            "attachments": attachments,
            "status": "sent",
        }

    def reply_message(self, config: object, message_id: str, payload: object, *, reply_all: bool = False) -> object:
        self.reply_calls.append((config, message_id, payload, reply_all))
        attachments = [
            {
                "id": item.get("id") or f"reply-att-{index}",
                "name": item["name"],
                "content_type": item.get("content_type", "application/octet-stream"),
                "size": len(base64.b64decode(item["content_base64"])),
                "is_inline": bool(item.get("is_inline", False)),
                "content_base64": item["content_base64"],
                "content_id": item.get("content_id", ""),
            }
            for index, item in enumerate(payload.get("attachments", []), start=1)
        ]
        return {
            "method": getattr(config, "default_method", "graph_api"),
            "message_id": "reply-001",
            "subject": "Re: Detail subject",
            "sender": getattr(config, "email", ""),
            "sender_name": "Sender",
            "received_at": "2026-03-13T00:00:00Z",
            "is_read": True,
            "has_attachments": bool(payload.get("attachments")),
            "preview": payload.get("body_text", "Reply body"),
            "source": "Graph API",
            "folder": "sentitems" if payload.get("send_now", True) else "drafts",
            "internet_message_id": "<reply-001>",
            "is_flagged": False,
            "importance": "normal",
            "conversation_id": "conv-1",
            "body_text": payload.get("body_text", "Reply body"),
            "body_html": payload.get("body_html"),
            "to": payload.get("to_recipients", ["sender@example.com"]),
            "cc": payload.get("cc_recipients", []),
            "bcc": payload.get("bcc_recipients", []),
            "headers": {"In-Reply-To": message_id},
            "attachments": attachments,
            "status": "sent" if payload.get("send_now", True) else "draft_saved",
        }

    def forward_message(self, config: object, message_id: str, payload: object) -> object:
        self.forward_calls.append((config, message_id, payload))
        attachments = [
            {
                "id": item.get("id") or f"forward-att-{index}",
                "name": item["name"],
                "content_type": item.get("content_type", "application/octet-stream"),
                "size": len(base64.b64decode(item["content_base64"])),
                "is_inline": bool(item.get("is_inline", False)),
                "content_base64": item["content_base64"],
                "content_id": item.get("content_id", ""),
            }
            for index, item in enumerate(payload.get("attachments", []), start=1)
        ]
        return {
            "method": getattr(config, "default_method", "graph_api"),
            "message_id": "forward-001",
            "subject": "Fwd: Detail subject",
            "sender": getattr(config, "email", ""),
            "sender_name": "Sender",
            "received_at": "2026-03-13T00:00:00Z",
            "is_read": True,
            "has_attachments": bool(payload.get("attachments")),
            "preview": payload.get("body_text", "Forward body"),
            "source": "Graph API",
            "folder": "sentitems",
            "internet_message_id": "<forward-001>",
            "is_flagged": False,
            "importance": "normal",
            "conversation_id": "conv-forward",
            "body_text": payload.get("body_text", "Forward body"),
            "body_html": payload.get("body_html"),
            "to": payload.get("to_recipients", ["target@example.com"]),
            "cc": payload.get("cc_recipients", []),
            "bcc": payload.get("bcc_recipients", []),
            "headers": {"References": message_id},
            "attachments": attachments,
            "status": "sent",
        }

    def upload_attachment(self, config: object, message_id: str, payload: object) -> object:
        self.upload_attachment_calls.append((config, message_id, payload))
        return {
            "id": "att-upload-001",
            "name": payload["name"],
            "content_type": payload.get("content_type", "text/plain"),
            "size": len(base64.b64decode(payload["content_base64"])),
            "is_inline": bool(payload.get("is_inline", False)),
            "content_base64": payload["content_base64"],
            "content_id": payload.get("content_id", ""),
        }

    def download_attachment(self, config: object, message_id: str, attachment_id: str) -> object:
        self.download_attachment_calls.append((config, message_id, attachment_id))
        return {
            "id": attachment_id,
            "name": "report.txt",
            "content_type": "text/plain",
            "size": 5,
            "is_inline": False,
            "content_base64": base64.b64encode(b"hello").decode("ascii"),
            "content_id": "",
        }

    def create_folder(self, config: object, payload: object) -> object:
        self.folder_create_calls.append((config, payload))
        return {
            "id": "Projects",
            "name": "Projects",
            "display_name": payload["display_name"],
            "kind": "custom",
            "total": 0,
            "unread": 0,
            "is_default": False,
            "status": "created",
        }

    def rename_folder(self, config: object, payload: object) -> object:
        self.folder_rename_calls.append((config, payload))
        return {
            "id": payload["display_name"],
            "name": payload["display_name"],
            "display_name": payload["display_name"],
            "kind": "custom",
            "total": 0,
            "unread": 0,
            "is_default": False,
            "status": "renamed",
        }

    def delete_folder(self, config: object, payload: object) -> object:
        self.folder_delete_calls.append((config, payload))
        return {"id": payload["folder_id"], "status": "deleted"}

    def resolve_mailbox_email(self, *, client_id: str, refresh_token: str, proxy: str | None = None) -> str:
        if self.resolve_email_error:
            raise MailboxError(self.resolve_email_error)
        resolved = self.resolved_emails.get((client_id, refresh_token))
        if not resolved:
            raise MailboxError("无法从授权信息中解析邮箱账号")
        return resolved


def build_client(
    store: MailboxStore,
    manager: RecordingManager,
    *,
    public_api_key: str | None = None,
    authenticate_admin: bool = True,
):
    app = create_app(manager=manager, store=store, public_api_key=public_api_key)
    app.config["TESTING"] = True
    client = app.test_client()
    if authenticate_admin:
        with client.session_transaction() as session:
            session["admin_authenticated"] = True
            session["admin_username"] = "admin"
    return client


def create_mailbox(
    store: MailboxStore,
    *,
    label: str,
    email: str,
    preferred_method: str = "graph_api",
    notes: str = "",
) -> object:
    return store.create_mailbox(
        {
            "label": label,
            "email": email,
            "client_id": f"{label}-client-id",
            "refresh_token": f"{label}-refresh-token",
            "proxy": "http://127.0.0.1:8080",
            "preferred_method": preferred_method,
            "notes": notes,
        }
    )


def test_change_admin_password_updates_login_password_and_persists(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    auth_path = tmp_path / "admin-auth.json"
    app = create_app(
        manager=manager,
        store=store,
        admin_password="admin123456",
        admin_auth_path=auth_path,
    )
    app.config["TESTING"] = True
    client = app.test_client()
    with client.session_transaction() as session:
        session["admin_authenticated"] = True
        session["admin_username"] = "admin"

    response = client.post(
        "/api/auth/password",
        json={
            "current_password": "admin123456",
            "new_password": "new-admin-123",
            "confirm_password": "new-admin-123",
        },
    )

    assert response.status_code == 200
    assert response.get_json() == {"updated": True, "username": "admin"}
    assert auth_path.exists()

    relogin_app = create_app(
        manager=manager,
        store=store,
        admin_password="admin123456",
        admin_auth_path=auth_path,
    )
    relogin_app.config["TESTING"] = True
    relogin_client = relogin_app.test_client()

    old_password_response = relogin_client.post(
        "/api/auth/login",
        json={"username": "admin", "password": "admin123456"},
    )
    assert old_password_response.status_code == 401

    new_password_response = relogin_client.post(
        "/api/auth/login",
        json={"username": "admin", "password": "new-admin-123"},
    )
    assert new_password_response.status_code == 200
    assert new_password_response.get_json() == {"authenticated": True, "username": "admin"}


def test_change_admin_password_rejects_invalid_current_password(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    app = create_app(manager=manager, store=store, admin_password="admin123456")
    app.config["TESTING"] = True
    client = app.test_client()
    with client.session_transaction() as session:
        session["admin_authenticated"] = True
        session["admin_username"] = "admin"

    response = client.post(
        "/api/auth/password",
        json={
            "current_password": "wrong-password",
            "new_password": "new-admin-123",
            "confirm_password": "new-admin-123",
        },
    )

    assert response.status_code == 400
    payload = response.get_json()
    assert payload["error"]["code"] == "invalid_current_password"


def test_change_admin_password_requires_minimum_length(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    app = create_app(manager=manager, store=store, admin_password="admin123456")
    app.config["TESTING"] = True
    client = app.test_client()
    with client.session_transaction() as session:
        session["admin_authenticated"] = True
        session["admin_username"] = "admin"

    response = client.post(
        "/api/auth/password",
        json={
            "current_password": "admin123456",
            "new_password": "short",
            "confirm_password": "short",
        },
    )

    assert response.status_code == 400
    payload = response.get_json()
    assert payload["error"]["code"] == "invalid_new_password"


def test_list_mailboxes_returns_summary_and_meta(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    create_mailbox(store, label="Alpha Box", email="alpha@example.com", notes="Project Red")
    create_mailbox(store, label="Beta Box", email="beta@example.com", notes="Project Blue")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.get("/api/mailboxes?q=BLUE&page=1&page_size=1")

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["meta"] == {
        "q": "BLUE",
        "page": 1,
        "page_size": 1,
        "total": 1,
        "total_pages": 1,
        "has_prev": False,
        "has_next": False,
    }
    assert len(payload["items"]) == 1
    item = payload["items"][0]
    assert item["label"] == "Beta Box"
    assert item["email"] == "beta@example.com"
    assert item["preferred_method"] == "graph_api"
    assert item["notes"] == "Project Blue"
    assert "client_id" not in item
    assert "refresh_token" not in item
    assert "proxy" not in item


def test_list_mailboxes_rejects_invalid_pagination(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.get("/api/mailboxes?page=0&page_size=101")

    assert response.status_code == 400
    payload = response.get_json()
    assert payload["error"]["message"] == "page 必须大于等于 1"


def test_mailbox_detail_still_returns_full_fields(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox = create_mailbox(
        store,
        label="Gamma Box",
        email="gamma@example.com",
        preferred_method="imap_new",
        notes="Detail check",
    )
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.get(f"/api/mailboxes/{mailbox.id}")

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["mailbox"]["client_id"] == "Gamma Box-client-id"
    assert payload["mailbox"]["refresh_token"] == "Gamma Box-refresh-token"
    assert payload["mailbox"]["proxy"] == "http://127.0.0.1:8080"


def test_test_connection_success(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/test-connection",
        json={
            "email": "demo@example.com",
            "client_id": "client-id",
            "refresh_token": "refresh-token",
            "proxy": "http://127.0.0.1:8080",
            "preferred_method": "graph_api",
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload == {
        "success": True,
        "method": "graph_api",
        "label": "Graph API",
        "message": "连接成功，最近邮件：Welcome mail",
    }
    config, request = manager.calls[0]
    assert config.email == "demo@example.com"
    assert config.client_id == "client-id"
    assert config.refresh_token == "refresh-token"
    assert config.proxy == "http://127.0.0.1:8080"
    assert config.default_method == "graph_api"
    assert request.method == "graph_api"
    assert request.top == 1


def test_test_connection_requires_required_fields(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/test-connection",
        json={
            "email": "demo@example.com",
            "preferred_method": "graph_api",
        },
    )

    assert response.status_code == 400
    payload = response.get_json()
    assert payload["error"]["message"] == "Client ID 不能为空"
    assert manager.calls == []


def test_test_connection_propagates_mailbox_error(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager(fail_emails={"demo@example.com"})
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/test-connection",
        json={
            "email": "demo@example.com",
            "client_id": "client-id",
            "refresh_token": "refresh-token",
            "preferred_method": "imap_new",
        },
    )

    assert response.status_code == 400
    payload = response.get_json()
    assert payload["error"]["message"] == "连接失败"


def test_batch_test_connection_returns_partial_failures(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox_a = create_mailbox(store, label="Mail A", email="a@example.com", preferred_method="graph_api")
    mailbox_b = create_mailbox(store, label="Mail B", email="b@example.com", preferred_method="imap_old")
    manager = RecordingManager(fail_emails={"b@example.com"})
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/test-connection/batch",
        json={
            "mailbox_ids": [mailbox_a.id, mailbox_b.id, 9999],
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["summary"] == {
        "processed": 3,
        "succeeded": 1,
        "failed": 2,
    }
    assert payload["results"][0]["success"] is True
    assert payload["results"][0]["method"] == "graph_api"
    assert payload["results"][1]["success"] is False
    assert payload["results"][1]["message"] == "连接失败"
    assert payload["results"][2]["success"] is False
    assert payload["results"][2]["message"] == "邮箱档案不存在"


def test_batch_test_connection_rejects_invalid_ids(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/test-connection/batch",
        json={
            "mailbox_ids": "not-an-array",
        },
    )

    assert response.status_code == 400
    payload = response.get_json()
    assert payload["error"]["message"] == "mailbox_ids 必须是数组"


def test_batch_update_preferred_method_returns_partial_failures(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox_a = create_mailbox(store, label="Mail C", email="c@example.com", preferred_method="graph_api")
    mailbox_b = create_mailbox(store, label="Mail D", email="d@example.com", preferred_method="imap_old")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/preferred-method/batch",
        json={
            "mailbox_ids": [mailbox_a.id, mailbox_b.id, 9999],
            "preferred_method": "imap_new",
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["summary"] == {
        "processed": 3,
        "succeeded": 2,
        "failed": 1,
    }
    assert payload["results"][0]["success"] is True
    assert payload["results"][0]["preferred_method"] == "imap_new"
    assert payload["results"][1]["success"] is True
    assert payload["results"][2]["success"] is False
    assert payload["results"][2]["message"] == "邮箱档案不存在"
    assert store.get_mailbox(mailbox_a.id).preferred_method == "imap_new"
    assert store.get_mailbox(mailbox_b.id).preferred_method == "imap_new"


def test_batch_update_preferred_method_requires_method(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox = create_mailbox(store, label="Mail E", email="e@example.com", preferred_method="graph_api")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/preferred-method/batch",
        json={
            "mailbox_ids": [mailbox.id],
        },
    )

    assert response.status_code == 400
    payload = response.get_json()
    assert payload["error"]["message"] == "请选择邮件接入方式"


def test_batch_delete_mailboxes_returns_partial_failures(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox_a = create_mailbox(store, label="Mail F", email="f@example.com", preferred_method="graph_api")
    mailbox_b = create_mailbox(store, label="Mail G", email="g@example.com", preferred_method="imap_new")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/delete/batch",
        json={
            "mailbox_ids": [mailbox_a.id, mailbox_b.id, 9999],
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["summary"] == {
        "processed": 3,
        "succeeded": 2,
        "failed": 1,
    }
    assert payload["results"][0]["success"] is True
    assert payload["results"][0]["message"] == "邮箱档案已删除"
    assert payload["results"][1]["success"] is True
    assert payload["results"][2]["success"] is False
    assert payload["results"][2]["message"] == "邮箱档案不存在"
    assert store.get_mailbox(mailbox_a.id) is None
    assert store.get_mailbox(mailbox_b.id) is None


def test_import_mailboxes_accepts_legacy_delimited_text(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/import",
        json={
            "raw_text": "legacy@example.com----reserved----legacy-client----legacy-token",
            "preferred_method": "graph_api",
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["summary"] == {
        "processed": 1,
        "created": 1,
        "updated": 0,
        "deduplicated": 0,
    }
    mailbox = store.get_mailbox_by_email("legacy@example.com")
    assert mailbox is not None
    assert mailbox.client_id == "legacy-client"
    assert mailbox.refresh_token == "legacy-token"
    assert mailbox.preferred_method == "graph_api"


def test_import_mailboxes_accepts_compact_delimited_text(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/import",
        json={
            "raw_text": "compact@example.com----compact-client----compact-token",
            "preferred_method": "imap_new",
        },
    )

    assert response.status_code == 200
    mailbox = store.get_mailbox_by_email("compact@example.com")
    assert mailbox is not None
    assert mailbox.client_id == "compact-client"
    assert mailbox.refresh_token == "compact-token"
    assert mailbox.preferred_method == "imap_new"


def test_import_mailboxes_accepts_reordered_delimited_text(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/import",
        json={
            "raw_text": (
                "11111111-1111-1111-1111-111111111111"
                "----M.C550_BAY.0.U.-example_refresh_token_value!"
                "----reorder@example.com"
            ),
            "preferred_method": "graph_api",
        },
    )

    assert response.status_code == 200
    mailbox = store.get_mailbox_by_email("reorder@example.com")
    assert mailbox is not None
    assert mailbox.client_id == "11111111-1111-1111-1111-111111111111"
    assert mailbox.refresh_token == "M.C550_BAY.0.U.-example_refresh_token_value!"


def test_import_mailboxes_accepts_keyed_delimited_text(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/import",
        json={
            "raw_text": (
                "rt=M.C550_BAY.0.U.-keyed_refresh_token!"
                "----email=keyed@example.com"
                "----clinetid=22222222-2222-2222-2222-222222222222"
                "----method=imap_old"
            ),
            "preferred_method": "graph_api",
        },
    )

    assert response.status_code == 200
    mailbox = store.get_mailbox_by_email("keyed@example.com")
    assert mailbox is not None
    assert mailbox.client_id == "22222222-2222-2222-2222-222222222222"
    assert mailbox.refresh_token == "M.C550_BAY.0.U.-keyed_refresh_token!"
    assert mailbox.preferred_method == "imap_old"


def test_import_mailboxes_accepts_csv_with_header(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/import",
        json={
            "raw_text": (
                "邮箱账号,Client ID,Refresh Token,默认方法\n"
                "csv@example.com,csv-client,csv-token,imap_old"
            ),
            "preferred_method": "graph_api",
        },
    )

    assert response.status_code == 200
    mailbox = store.get_mailbox_by_email("csv@example.com")
    assert mailbox is not None
    assert mailbox.client_id == "csv-client"
    assert mailbox.refresh_token == "csv-token"
    assert mailbox.preferred_method == "imap_old"


def test_import_mailboxes_accepts_csv_without_email_column(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager(
        resolved_emails={
            (
                "33333333-3333-3333-3333-333333333333",
                "M.C550_BAY.0.U.-csv_missing_email!",
            ): "csv-resolved@example.com"
        }
    )
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/import",
        json={
            "raw_text": (
                "client_id,refresh_token\n"
                "33333333-3333-3333-3333-333333333333,M.C550_BAY.0.U.-csv_missing_email!"
            ),
            "preferred_method": "graph_api",
        },
    )

    assert response.status_code == 200
    mailbox = store.get_mailbox_by_email("csv-resolved@example.com")
    assert mailbox is not None
    assert mailbox.client_id == "33333333-3333-3333-3333-333333333333"
    assert mailbox.refresh_token == "M.C550_BAY.0.U.-csv_missing_email!"


def test_import_mailboxes_accepts_json_array(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/import",
        json={
            "raw_text": (
                '[{"email":"json-a@example.com","client_id":"json-client-a","refresh_token":"json-token-a"},'
                '{"email":"json-b@example.com","client_id":"json-client-b","refresh_token":"json-token-b","preferred_method":"imap_new"}]'
            ),
            "preferred_method": "graph_api",
        },
    )

    assert response.status_code == 200
    mailbox_a = store.get_mailbox_by_email("json-a@example.com")
    mailbox_b = store.get_mailbox_by_email("json-b@example.com")
    assert mailbox_a is not None
    assert mailbox_b is not None
    assert mailbox_a.preferred_method == "graph_api"
    assert mailbox_b.preferred_method == "imap_new"


def test_import_mailboxes_rejects_invalid_tabular_row(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/import",
        json={
            "raw_text": "broken@example.com,missing-token",
            "preferred_method": "graph_api",
        },
    )

    assert response.status_code == 400
    payload = response.get_json()
    assert "CSV/TSV" in payload["error"]["message"]


def test_import_mailboxes_resolves_email_when_missing(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager(
        resolved_emails={
            (
                "11111111-1111-1111-1111-111111111111",
                "M.C550_BAY.0.U.-missing_email!",
            ): "resolved@example.com"
        }
    )
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/import",
        json={
            "raw_text": "11111111-1111-1111-1111-111111111111----M.C550_BAY.0.U.-missing_email!",
            "preferred_method": "graph_api",
        },
    )

    assert response.status_code == 200
    mailbox = store.get_mailbox_by_email("resolved@example.com")
    assert mailbox is not None
    assert mailbox.client_id == "11111111-1111-1111-1111-111111111111"
    assert mailbox.refresh_token == "M.C550_BAY.0.U.-missing_email!"


def test_import_mailboxes_rejects_when_missing_email_cannot_be_resolved(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager(resolve_email_error="无法自动解析邮箱账号")
    client = build_client(store, manager)

    response = client.post(
        "/api/mailboxes/import",
        json={
            "raw_text": "11111111-1111-1111-1111-111111111111----M.C550_BAY.0.U.-missing_email!",
            "preferred_method": "graph_api",
        },
    )

    assert response.status_code == 400
    payload = response.get_json()
    assert "无法自动解析邮箱账号" in payload["error"]["message"]


def test_key_mailbox_messages_returns_public_mailbox_and_messages(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox = create_mailbox(
        store,
        label="Public Box",
        email="public@example.com",
        preferred_method="graph_api",
        notes="Key access",
    )
    manager = RecordingManager()
    client = build_client(store, manager, public_api_key="secret-key")

    response = client.post(
        "/api/key/mailbox/messages",
        headers={"X-InboxOps-Key": "secret-key"},
        json={
            "email": "public@example.com",
            "method": "imap_new",
            "top": 5,
            "keyword": "Welcome",
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["mailbox"] == {
        "id": mailbox.id,
        "label": "Public Box",
        "email": "public@example.com",
        "preferred_method": "graph_api",
        "notes": "Key access",
        "created_at": mailbox.created_at,
        "updated_at": mailbox.updated_at,
    }
    assert payload["method"] == "imap_new"
    assert payload["count"] == 1
    assert payload["messages"][0]["subject"] == "Welcome mail"
    assert "client_id" not in payload["mailbox"]
    assert "refresh_token" not in payload["mailbox"]
    assert "proxy" not in payload["mailbox"]

    config, request = manager.calls[0]
    assert config.email == "public@example.com"
    assert config.client_id == "Public Box-client-id"
    assert config.refresh_token == "Public Box-refresh-token"
    assert request.method == "imap_new"
    assert request.top == 5
    assert request.keyword == "Welcome"


def test_key_mailbox_message_returns_detail_with_body_key(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox = create_mailbox(store, label="Detail Box", email="detail@example.com", preferred_method="imap_old")
    manager = RecordingManager()
    client = build_client(store, manager, public_api_key="secret-key")

    response = client.post(
        "/api/key/mailbox/message",
        json={
            "email": "detail@example.com",
            "message_id": "msg-001",
            "api_key": "secret-key",
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["mailbox"]["id"] == mailbox.id
    assert payload["mailbox"]["email"] == "detail@example.com"
    assert payload["message"]["message_id"] == "msg-001"
    assert payload["message"]["subject"] == "Detail subject"
    assert payload["message"]["body_text"] == "Hello detail"
    assert payload["message"]["to_recipients"] == ["demo@example.com"]
    assert "client_id" not in payload["mailbox"]
    assert "refresh_token" not in payload["mailbox"]
    assert "proxy" not in payload["mailbox"]

    config, request = manager.detail_calls[0]
    assert config.email == "detail@example.com"
    assert request.method == "imap_old"
    assert request.message_id == "msg-001"


def test_key_mailbox_messages_rejects_invalid_key(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    create_mailbox(store, label="Invalid Key Box", email="invalid@example.com")
    manager = RecordingManager()
    client = build_client(store, manager, public_api_key="secret-key", authenticate_admin=False)

    response = client.post(
        "/api/key/mailbox/messages",
        headers={"X-InboxOps-Key": "wrong-key"},
        json={"email": "invalid@example.com"},
    )

    assert response.status_code == 401
    payload = response.get_json()
    assert payload["error"]["message"] == "访问 Key 无效"


def test_key_mailbox_messages_requires_configured_key(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    create_mailbox(store, label="No Key Box", email="nokey@example.com")
    manager = RecordingManager()
    client = build_client(store, manager, public_api_key="", authenticate_admin=False)

    response = client.post(
        "/api/key/mailbox/messages",
        json={
            "email": "nokey@example.com",
            "key": "secret-key",
        },
    )

    assert response.status_code == 503
    payload = response.get_json()
    assert payload["error"]["message"] == "项目未配置访问 Key"


def test_key_mailbox_messages_returns_404_for_unknown_email(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    manager = RecordingManager()
    client = build_client(store, manager, public_api_key="secret-key", authenticate_admin=False)

    response = client.post(
        "/api/key/mailbox/messages",
        headers={"X-InboxOps-Key": "secret-key"},
        json={"email": "missing@example.com"},
    )

    assert response.status_code == 404
    payload = response.get_json()
    assert payload["error"]["message"] == "邮箱档案不存在"


def test_mailbox_folders_returns_folder_list(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox = create_mailbox(store, label="Folders Box", email="folders@example.com")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailbox/folders",
        json={
            "mailbox_id": mailbox.id,
            "method": "imap_new",
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["method"] == "imap_new"
    assert payload["folders"][0]["display_name"] == "收件箱"
    assert payload["folders"][1]["kind"] == "archive"
    config, method = manager.folder_calls[0]
    assert config.email == "folders@example.com"
    assert method == "imap_new"


def test_mailbox_messages_supports_folder_filters_and_meta(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox = create_mailbox(store, label="Messages Box", email="messages@example.com")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailbox/messages",
        json={
            "mailbox_id": mailbox.id,
            "method": "imap_old",
            "folder": "archive",
            "page": 2,
            "page_size": 25,
            "read_state": "unread",
            "has_attachments_only": True,
            "flagged_only": True,
            "importance": "high",
            "sort_order": "asc",
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["folder"] == "archive"
    assert payload["meta"]["page"] == 2
    assert payload["meta"]["page_size"] == 25
    assert payload["messages"][0]["folder"] == "archive"
    assert payload["messages"][0]["is_flagged"] is True
    assert payload["messages"][0]["importance"] == "high"

    config, request = manager.calls[0]
    assert config.email == "messages@example.com"
    assert request.folder == "archive"
    assert request.page == 2
    assert request.page_size == 25
    assert request.read_state == "unread"
    assert request.has_attachments_only is True
    assert request.flagged_only is True
    assert request.importance == "high"
    assert request.sort_order == "asc"


def test_mailbox_message_flag_state_returns_updated_message(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox = create_mailbox(store, label="Flag Box", email="flag@example.com")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailbox/message/flag-state",
        json={
            "mailbox_id": mailbox.id,
            "message_id": "msg-001",
            "folder": "INBOX",
            "is_flagged": True,
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["message"]["message_id"] == "msg-001"
    assert payload["message"]["is_flagged"] is True
    assert payload["message"]["body_html"] == "<p>Hello <strong>detail</strong></p>"
    config, request = manager.flag_calls[0]
    assert config.email == "flag@example.com"
    assert request.message_id == "msg-001"
    assert request.is_flagged is True
    assert request.folder == "INBOX"


def test_mailbox_message_move_and_delete_routes_return_results(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox = create_mailbox(store, label="Move Box", email="move@example.com")
    manager = RecordingManager()
    client = build_client(store, manager)

    move_response = client.post(
        "/api/mailbox/message/move",
        json={
            "mailbox_id": mailbox.id,
            "message_id": "msg-001",
            "folder": "INBOX",
            "destination_folder": "archive",
        },
    )

    delete_response = client.post(
        "/api/mailbox/message/delete",
        json={
            "mailbox_id": mailbox.id,
            "message_id": "msg-001",
            "folder": "archive",
        },
    )

    assert move_response.status_code == 200
    move_payload = move_response.get_json()
    assert move_payload["result"]["destination_folder"] == "archive"
    assert move_payload["result"]["status"] == "moved"

    assert delete_response.status_code == 200
    delete_payload = delete_response.get_json()
    assert delete_payload["result"]["folder"] == "archive"
    assert delete_payload["result"]["status"] == "deleted"

    _, move_request = manager.move_calls[0]
    assert move_request.destination_folder == "archive"
    _, delete_request = manager.delete_calls[0]
    assert delete_request.folder == "archive"


def test_mailbox_batch_actions_support_flag_move_and_delete(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox = create_mailbox(store, label="Batch Box", email="batch@example.com")
    manager = RecordingManager()
    client = build_client(store, manager)

    response = client.post(
        "/api/mailbox/messages/actions/batch",
        json={
            "mailbox_id": mailbox.id,
            "method": "graph_api",
            "folder": "INBOX",
            "message_ids": ["msg-001", "msg-002"],
            "action": "move",
            "destination_folder": "archive",
        },
    )

    assert response.status_code == 200
    payload = response.get_json()
    assert payload["summary"] == {
        "processed": 2,
        "succeeded": 2,
        "failed": 0,
    }
    assert payload["results"][0]["success"] is True
    assert payload["results"][0]["result"]["destination_folder"] == "archive"
    assert len(manager.move_calls) == 2


def test_mailbox_compose_routes_cover_draft_send_reply_and_forward(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox = create_mailbox(store, label="Compose Box", email="compose@example.com")
    manager = RecordingManager()
    client = build_client(store, manager)
    content = base64.b64encode(b"hello").decode("ascii")

    draft_response = client.post(
        "/api/mailbox/message/draft",
        json={
            "mailbox_id": mailbox.id,
            "subject": "Draft subject",
            "body_text": "Draft body",
            "to_recipients": ["draft@example.com"],
            "attachments": [{"name": "draft.txt", "content_base64": content, "content_type": "text/plain"}],
        },
    )
    send_response = client.post(
        "/api/mailbox/message/send",
        json={
            "mailbox_id": mailbox.id,
            "subject": "Send subject",
            "body_text": "Send body",
            "to_recipients": ["target@example.com"],
        },
    )
    reply_all_response = client.post(
        "/api/mailbox/message/reply-all",
        json={
            "mailbox_id": mailbox.id,
            "message_id": "msg-001",
            "body_text": "Reply body",
        },
    )
    forward_response = client.post(
        "/api/mailbox/message/forward",
        json={
            "mailbox_id": mailbox.id,
            "message_id": "msg-001",
            "body_text": "Forward body",
            "to_recipients": ["forward@example.com"],
        },
    )

    assert draft_response.status_code == 200
    assert draft_response.get_json()["message"]["folder"] == "drafts"
    assert send_response.status_code == 200
    assert send_response.get_json()["message"]["status"] == "sent"
    assert reply_all_response.status_code == 200
    assert reply_all_response.get_json()["message"]["conversation_id"] == "conv-1"
    assert forward_response.status_code == 200
    assert forward_response.get_json()["message"]["subject"] == "Fwd: Detail subject"
    assert len(manager.draft_calls) == 1
    assert len(manager.send_calls) == 1
    assert len(manager.reply_calls) == 1
    assert manager.reply_calls[0][3] is True
    assert len(manager.forward_calls) == 1


def test_mailbox_attachment_folder_meta_search_thread_and_rules_routes(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox = create_mailbox(store, label="Ops Box", email="ops@example.com")
    manager = RecordingManager()
    client = build_client(store, manager)
    content = base64.b64encode(b"hello").decode("ascii")

    upload_response = client.post(
        "/api/mailbox/message/attachment/upload",
        json={
            "mailbox_id": mailbox.id,
            "message_id": "draft-001",
            "name": "report.txt",
            "content_type": "text/plain",
            "content_base64": content,
        },
    )
    download_response = client.post(
        "/api/mailbox/message/attachment/download",
        json={
            "mailbox_id": mailbox.id,
            "message_id": "draft-001",
            "attachment_id": "att-upload-001",
        },
    )
    create_folder_response = client.post(
        "/api/mailbox/folder/create",
        json={"mailbox_id": mailbox.id, "display_name": "Projects"},
    )
    rename_folder_response = client.post(
        "/api/mailbox/folder/rename",
        json={"mailbox_id": mailbox.id, "folder_id": "Projects", "display_name": "Projects 2026"},
    )
    delete_folder_response = client.post(
        "/api/mailbox/folder/delete",
        json={"mailbox_id": mailbox.id, "folder_id": "Projects 2026"},
    )
    client.post("/api/mailbox/messages", json={"mailbox_id": mailbox.id, "method": "graph_api"})
    meta_response = client.post(
        "/api/mailbox/message/meta",
        json={
            "mailbox_id": mailbox.id,
            "method": "graph_api",
            "message_id": "msg-001",
            "tags": ["vip", "follow-up"],
            "notes": "Need action",
            "follow_up": "today",
            "snoozed_until": "2026-03-23T00:00:00Z",
            "status": "snoozed",
        },
    )
    search_response = client.post(
        "/api/mailboxes/messages/search",
        json={"query": "welcome", "mailbox_ids": [mailbox.id], "method": "graph_api"},
    )
    thread_response = client.post(
        "/api/mailbox/thread",
        json={"mailbox_id": mailbox.id, "method": "graph_api", "message_id": "msg-001"},
    )
    rule_response = client.post(
        "/api/mailbox/rules",
        json={
            "mailbox_id": mailbox.id,
            "name": "Read Welcome",
            "enabled": True,
            "priority": 10,
            "conditions": {"subject_contains": "Welcome"},
            "actions": {"mark_read": True, "tags": ["ruled"]},
        },
    )
    apply_response = client.post(
        "/api/mailbox/rules/apply",
        json={"mailbox_id": mailbox.id, "method": "graph_api"},
    )

    assert upload_response.status_code == 200
    assert upload_response.get_json()["attachment"]["name"] == "report.txt"
    assert download_response.status_code == 200
    assert download_response.get_json()["attachment"]["content_base64"] == content
    assert create_folder_response.status_code == 200
    assert rename_folder_response.status_code == 200
    assert delete_folder_response.status_code == 200
    assert meta_response.status_code == 200
    assert meta_response.get_json()["meta"]["tags"] == ["vip", "follow-up"]
    assert search_response.status_code == 200
    assert search_response.get_json()["items"][0]["subject"] == "Welcome mail"
    assert thread_response.status_code == 200
    assert thread_response.get_json()["count"] == 1
    assert rule_response.status_code == 201
    assert apply_response.status_code == 200
    assert apply_response.get_json()["results"][0]["matched"] == 1
    assert len(manager.read_calls) == 1
    assert len(manager.upload_attachment_calls) == 1
    assert len(manager.download_attachment_calls) == 1
    assert len(manager.folder_create_calls) == 1
    assert len(manager.folder_rename_calls) == 1
    assert len(manager.folder_delete_calls) == 1


def test_audit_logs_and_sync_center_routes_return_persisted_state(tmp_path) -> None:
    store = MailboxStore(tmp_path / "mailboxes.db")
    mailbox = create_mailbox(store, label="Sync Box", email="sync@example.com")
    manager = RecordingManager()
    client = build_client(store, manager)

    client.post(
        "/api/mailbox/rules",
        json={
            "mailbox_id": mailbox.id,
            "name": "Sync Rule",
            "enabled": True,
            "priority": 5,
            "conditions": {"subject_contains": "Welcome"},
            "actions": {"mark_read": True},
        },
    )
    sync_response = client.post(
        "/api/mailbox/sync/run",
        json={
            "mailbox_id": mailbox.id,
            "method": "graph_api",
            "folder_limit": 1,
            "message_limit": 1,
            "include_body": True,
            "apply_rules": False,
        },
    )
    status_response = client.post(
        "/api/mailbox/sync/status",
        json={"mailbox_id": mailbox.id, "method": "graph_api"},
    )
    audit_response = client.post(
        "/api/audit/logs",
        json={"mailbox_id": mailbox.id, "page": 1, "page_size": 20},
    )

    assert sync_response.status_code == 200
    sync_payload = sync_response.get_json()
    assert sync_payload["count"] == 1
    assert sync_payload["results"][0]["job"]["status"] == "completed"

    assert status_response.status_code == 200
    status_payload = status_response.get_json()
    assert status_payload["items"][0]["jobs"][0]["status"] == "completed"
    assert status_payload["items"][0]["states"][0]["status"] == "completed"

    assert audit_response.status_code == 200
    audit_payload = audit_response.get_json()
    assert audit_payload["meta"]["total"] >= 1
