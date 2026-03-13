from __future__ import annotations

from app import create_app
from services.outlook_manager import MailboxError, MessageDetail, MessageListResult, MessageSummary
from services.storage import MailboxStore


class RecordingManager:
    def __init__(self, *, fail_emails: set[str] | None = None) -> None:
        self.fail_emails = fail_emails or set()
        self.calls: list[tuple[object, object]] = []
        self.detail_calls: list[tuple[object, object]] = []

    def list_messages(self, config: object, request: object) -> object:
        self.calls.append((config, request))
        if getattr(config, "email", "") in self.fail_emails:
            raise MailboxError("连接失败")
        method = getattr(request, "method", "graph_api")
        return MessageListResult(
            method=method,
            total=1,
            returned=1,
            messages=[
                MessageSummary(
                    method=method,
                    message_id="msg-001",
                    subject="Welcome mail",
                    sender="sender@example.com",
                    sender_name="Sender",
                    received_at="2026-03-13T00:00:00Z",
                    is_read=False,
                    has_attachments=False,
                    preview="Preview",
                    source="Graph API",
                )
            ],
        )

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
            has_attachments=False,
            preview="Preview",
            source="Graph API",
            internet_message_id="<demo>",
            body_text="Hello detail",
            to_recipients=["demo@example.com"],
            cc_recipients=[],
            bcc_recipients=[],
            headers={},
            conversation_id="conv-1",
        )


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
