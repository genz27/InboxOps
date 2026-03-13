from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

DEFAULT_METHOD = "graph_api"
VALID_METHODS = {DEFAULT_METHOD, "imap_new", "imap_old"}


@dataclass(slots=True)
class MailboxProfile:
    id: int
    label: str
    email: str
    client_id: str
    refresh_token: str
    proxy: str | None
    preferred_method: str
    notes: str
    created_at: str
    updated_at: str

    def to_runtime_config(self) -> dict[str, Any]:
        """给主线程一个统一的运行时配置视图。"""
        return {
            "email": self.email,
            "client_id": self.client_id,
            "refresh_token": self.refresh_token,
            "proxy": self.proxy or "",
            "preferred_method": self.preferred_method,
        }


@dataclass(slots=True)
class MailboxSummary:
    id: int
    label: str
    email: str
    preferred_method: str
    notes: str
    created_at: str
    updated_at: str


class MailboxStoreError(RuntimeError):
    """多邮箱档案存储层的统一异常。"""


class MailboxStore:
    def __init__(self, db_path: str | Path = "data/mailboxes.db") -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize_schema()

    def list_mailboxes(self) -> list[MailboxProfile]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT
                    id,
                    label,
                    email,
                    client_id,
                    refresh_token,
                    proxy,
                    preferred_method,
                    notes,
                    created_at,
                    updated_at
                FROM mailboxes
                ORDER BY updated_at DESC, id DESC
                """
            ).fetchall()
        return [self._row_to_profile(row) for row in rows]

    def search_mailboxes_summary(
        self,
        query: str | None,
        *,
        page: int,
        page_size: int,
    ) -> tuple[list[MailboxSummary], int]:
        normalized_query = self._normalize_text(query) or ""
        where_clause = ""
        parameters: list[Any] = []

        if normalized_query:
            pattern = f"%{normalized_query.lower()}%"
            where_clause = """
                WHERE
                    lower(label) LIKE ?
                    OR lower(email) LIKE ?
                    OR lower(notes) LIKE ?
            """
            parameters.extend([pattern, pattern, pattern])

        count_sql = f"""
            SELECT COUNT(*) AS total
            FROM mailboxes
            {where_clause}
        """
        query_sql = f"""
            SELECT
                id,
                label,
                email,
                preferred_method,
                notes,
                created_at,
                updated_at
            FROM mailboxes
            {where_clause}
            ORDER BY updated_at DESC, id DESC
            LIMIT ? OFFSET ?
        """
        offset = (page - 1) * page_size

        with self._connect() as connection:
            total = int(connection.execute(count_sql, parameters).fetchone()["total"])
            rows = connection.execute(
                query_sql,
                [*parameters, page_size, offset],
            ).fetchall()

        return [self._row_to_summary(row) for row in rows], total

    def get_mailbox(self, mailbox_id: int) -> MailboxProfile | None:
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT
                    id,
                    label,
                    email,
                    client_id,
                    refresh_token,
                    proxy,
                    preferred_method,
                    notes,
                    created_at,
                    updated_at
                FROM mailboxes
                WHERE id = ?
                """,
                (mailbox_id,),
            ).fetchone()
        return self._row_to_profile(row) if row else None

    def create_mailbox(self, payload: dict[str, Any]) -> MailboxProfile:
        values = self._normalize_payload(payload, partial=False)
        now = self._utc_now()

        with self._connect() as connection:
            try:
                cursor = connection.execute(
                    """
                    INSERT INTO mailboxes (
                        label,
                        email,
                        client_id,
                        refresh_token,
                        proxy,
                        preferred_method,
                        notes,
                        created_at,
                        updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        values["label"],
                        values["email"],
                        values["client_id"],
                        values["refresh_token"],
                        values["proxy"],
                        values["preferred_method"],
                        values["notes"],
                        now,
                        now,
                    ),
                )
            except sqlite3.IntegrityError as exc:
                raise MailboxStoreError("该邮箱账号已存在，不能重复添加") from exc

        created = self.get_mailbox(int(cursor.lastrowid))
        if not created:
            raise MailboxStoreError("邮箱档案创建后读取失败")
        return created

    def get_mailbox_by_email(self, email: str) -> MailboxProfile | None:
        normalized_email = self._normalize_text(email)
        if not normalized_email:
            return None

        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT
                    id,
                    label,
                    email,
                    client_id,
                    refresh_token,
                    proxy,
                    preferred_method,
                    notes,
                    created_at,
                    updated_at
                FROM mailboxes
                WHERE lower(email) = lower(?)
                LIMIT 1
                """,
                (normalized_email,),
            ).fetchone()
        return self._row_to_profile(row) if row else None

    def import_mailboxes(self, payloads: list[dict[str, Any]]) -> tuple[dict[str, int], list[MailboxProfile]]:
        if not isinstance(payloads, list):
            raise MailboxStoreError("批量导入参数必须是数组")

        deduplicated_payloads: dict[str, dict[str, str]] = {}
        processed = 0

        for payload in payloads:
            if not isinstance(payload, dict):
                raise MailboxStoreError("批量导入的每一项都必须是对象")

            email = self._normalize_text(payload.get("email"))
            client_id = self._normalize_text(payload.get("client_id"))
            refresh_token = self._normalize_text(payload.get("refresh_token"))
            preferred_method = self._normalize_text(payload.get("preferred_method"), fallback=DEFAULT_METHOD) or DEFAULT_METHOD

            if not email:
                raise MailboxStoreError("批量导入缺少邮箱账号")
            if not client_id:
                raise MailboxStoreError(f"{email} 缺少 Client ID")
            if not refresh_token:
                raise MailboxStoreError(f"{email} 缺少 Refresh Token")
            if preferred_method not in VALID_METHODS:
                raise MailboxStoreError("preferred_method 仅支持 graph_api、imap_new、imap_old")

            processed += 1
            dedupe_key = email.casefold()
            if dedupe_key in deduplicated_payloads:
                deduplicated_payloads.pop(dedupe_key)
            deduplicated_payloads[dedupe_key] = {
                "email": email,
                "client_id": client_id,
                "refresh_token": refresh_token,
                "preferred_method": preferred_method,
            }

        if not deduplicated_payloads:
            raise MailboxStoreError("批量导入内容不能为空")

        created = 0
        updated = 0
        imported_mailboxes: list[MailboxProfile] = []

        for values in deduplicated_payloads.values():
            existing = self.get_mailbox_by_email(values["email"])
            if existing:
                mailbox = self.update_mailbox(
                    existing.id,
                    {
                        "client_id": values["client_id"],
                        "refresh_token": values["refresh_token"],
                        "preferred_method": values["preferred_method"],
                    },
                )
                updated += 1
            else:
                mailbox = self.create_mailbox(
                    {
                        "label": values["email"],
                        "email": values["email"],
                        "client_id": values["client_id"],
                        "refresh_token": values["refresh_token"],
                        "proxy": "",
                        "preferred_method": values["preferred_method"],
                        "notes": "",
                    }
                )
                created += 1
            imported_mailboxes.append(mailbox)

        summary = {
            "processed": processed,
            "created": created,
            "updated": updated,
            "deduplicated": processed - len(deduplicated_payloads),
        }
        return summary, imported_mailboxes

    def update_mailbox(self, mailbox_id: int, payload: dict[str, Any]) -> MailboxProfile:
        existing = self.get_mailbox(mailbox_id)
        if not existing:
            raise MailboxStoreError("要更新的邮箱档案不存在")

        values = self._normalize_payload(payload, partial=True, current=existing)
        now = self._utc_now()

        with self._connect() as connection:
            try:
                connection.execute(
                    """
                    UPDATE mailboxes
                    SET
                        label = ?,
                        email = ?,
                        client_id = ?,
                        refresh_token = ?,
                        proxy = ?,
                        preferred_method = ?,
                        notes = ?,
                        updated_at = ?
                    WHERE id = ?
                    """,
                    (
                        values["label"],
                        values["email"],
                        values["client_id"],
                        values["refresh_token"],
                        values["proxy"],
                        values["preferred_method"],
                        values["notes"],
                        now,
                        mailbox_id,
                    ),
                )
            except sqlite3.IntegrityError as exc:
                raise MailboxStoreError("更新失败：邮箱账号与其他档案重复") from exc

        updated = self.get_mailbox(mailbox_id)
        if not updated:
            raise MailboxStoreError("邮箱档案更新后读取失败")
        return updated

    def delete_mailbox(self, mailbox_id: int) -> bool:
        with self._connect() as connection:
            cursor = connection.execute("DELETE FROM mailboxes WHERE id = ?", (mailbox_id,))
        return cursor.rowcount > 0

    def _initialize_schema(self) -> None:
        # 初始化时自动建表，主线程不需要额外执行迁移脚本。
        with self._connect() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS mailboxes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    label TEXT NOT NULL,
                    email TEXT NOT NULL UNIQUE,
                    client_id TEXT NOT NULL,
                    refresh_token TEXT NOT NULL,
                    proxy TEXT DEFAULT '',
                    preferred_method TEXT NOT NULL DEFAULT 'graph_api',
                    notes TEXT DEFAULT '',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                )
                """
            )

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        return connection

    def _normalize_payload(
        self,
        payload: dict[str, Any],
        *,
        partial: bool,
        current: MailboxProfile | None = None,
    ) -> dict[str, str]:
        if not isinstance(payload, dict):
            raise MailboxStoreError("邮箱档案参数必须是对象")

        # 新建和更新共用一套归一化逻辑，确保数据库字段始终一致。
        email = self._normalize_text(payload.get("email"), fallback=current.email if current else None)
        client_id = self._normalize_text(
            payload.get("client_id"),
            fallback=current.client_id if current else None,
        )
        refresh_token = self._normalize_text(
            payload.get("refresh_token"),
            fallback=current.refresh_token if current else None,
        )

        if not partial and not email:
            raise MailboxStoreError("邮箱账号不能为空")
        if not partial and not client_id:
            raise MailboxStoreError("Client ID 不能为空")
        if not partial and not refresh_token:
            raise MailboxStoreError("Refresh Token 不能为空")

        label = self._normalize_text(payload.get("label"), fallback=current.label if current else None)
        label = label or email
        proxy = self._normalize_text(payload.get("proxy"), fallback=current.proxy if current else None) or ""
        notes = self._normalize_text(payload.get("notes"), fallback=current.notes if current else None) or ""
        preferred_method = (
            self._normalize_text(
                payload.get("preferred_method"),
                fallback=current.preferred_method if current else DEFAULT_METHOD,
            )
            or DEFAULT_METHOD
        )

        if preferred_method not in VALID_METHODS:
            raise MailboxStoreError("preferred_method 仅支持 graph_api、imap_new、imap_old")

        return {
            "label": label,
            "email": email or "",
            "client_id": client_id or "",
            "refresh_token": refresh_token or "",
            "proxy": proxy,
            "preferred_method": preferred_method,
            "notes": notes,
        }

    @staticmethod
    def _normalize_text(value: Any, fallback: str | None = None) -> str | None:
        if value is None:
            return fallback
        if not isinstance(value, str):
            raise MailboxStoreError("邮箱档案字段必须是字符串")
        cleaned = value.strip()
        return cleaned or None

    @staticmethod
    def _row_to_profile(row: sqlite3.Row) -> MailboxProfile:
        return MailboxProfile(
            id=int(row["id"]),
            label=str(row["label"]),
            email=str(row["email"]),
            client_id=str(row["client_id"]),
            refresh_token=str(row["refresh_token"]),
            proxy=str(row["proxy"]) if row["proxy"] else None,
            preferred_method=str(row["preferred_method"]),
            notes=str(row["notes"] or ""),
            created_at=str(row["created_at"]),
            updated_at=str(row["updated_at"]),
        )

    @staticmethod
    def _row_to_summary(row: sqlite3.Row) -> MailboxSummary:
        return MailboxSummary(
            id=int(row["id"]),
            label=str(row["label"]),
            email=str(row["email"]),
            preferred_method=str(row["preferred_method"]),
            notes=str(row["notes"] or ""),
            created_at=str(row["created_at"]),
            updated_at=str(row["updated_at"]),
        )

    @staticmethod
    def _utc_now() -> str:
        return datetime.now(UTC).isoformat(timespec="seconds").replace("+00:00", "Z")



