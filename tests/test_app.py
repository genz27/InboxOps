from __future__ import annotations

from types import SimpleNamespace

from app import create_app
from services.outlook_manager import MailboxError
from services.storage import MailboxStore


class RecordingManager:
    def __init__(self, *, fail_emails: set[str] | None = None) -> None:
        self.fail_emails = fail_emails or set()
        self.calls: list[tuple[object, object]] = []

    def list_messages(self, config: object, request: object) -> object:
        self.calls.append((config, request))
        if getattr(config, "email", "") in self.fail_emails:
            raise MailboxError("连接失败")
        return SimpleNamespace(messages=[SimpleNamespace(subject="Welcome mail")])


def build_client(store: MailboxStore, manager: RecordingManager):
    app = create_app(manager=manager, store=store)
    app.config["TESTING"] = True
    client = app.test_client()
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
