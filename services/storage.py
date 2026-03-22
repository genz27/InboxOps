from __future__ import annotations

import json
import sqlite3
import uuid
from dataclasses import asdict, dataclass, field, is_dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, Iterable

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


@dataclass(slots=True)
class FolderCacheEntry:
    mailbox_id: int
    method: str
    folder_id: str
    name: str
    display_name: str
    kind: str
    total: int
    unread: int
    is_default: bool
    cached_at: str
    updated_at: str


@dataclass(slots=True)
class CachedAttachment:
    mailbox_id: int
    method: str
    message_id: str
    attachment_id: str
    name: str
    content_type: str
    size: int
    is_inline: bool
    content_base64: str = ""
    cached_at: str = ""
    updated_at: str = ""


@dataclass(slots=True)
class MessageMeta:
    mailbox_id: int
    method: str
    message_id: str
    tags: list[str] = field(default_factory=list)
    follow_up: str = ""
    notes: str = ""
    snoozed_until: str = ""
    status: str = "active"
    updated_at: str = ""


@dataclass(slots=True)
class CachedMessage:
    mailbox_id: int
    mailbox_label: str
    mailbox_email: str
    method: str
    message_id: str
    subject: str
    sender: str
    sender_name: str
    received_at: str
    is_read: bool
    is_flagged: bool
    importance: str
    has_attachments: bool
    preview: str
    body_text: str
    body_html: str | None
    folder: str
    internet_message_id: str
    conversation_id: str
    to_recipients: list[str] = field(default_factory=list)
    cc_recipients: list[str] = field(default_factory=list)
    bcc_recipients: list[str] = field(default_factory=list)
    headers: dict[str, str] = field(default_factory=dict)
    in_reply_to: str = ""
    references: list[str] = field(default_factory=list)
    attachments: list[CachedAttachment] = field(default_factory=list)
    meta: MessageMeta | None = None
    cached_at: str = ""
    updated_at: str = ""


@dataclass(slots=True)
class MessageSearchResult:
    mailbox_id: int
    mailbox_label: str
    mailbox_email: str
    method: str
    message_id: str
    subject: str
    sender: str
    sender_name: str
    received_at: str
    is_read: bool
    is_flagged: bool
    importance: str
    has_attachments: bool
    preview: str
    folder: str
    internet_message_id: str
    conversation_id: str
    meta: MessageMeta | None = None


@dataclass(slots=True)
class SavedRule:
    id: int
    mailbox_id: int
    name: str
    enabled: bool
    priority: int
    conditions: dict[str, Any]
    actions: dict[str, Any]
    created_at: str
    updated_at: str


@dataclass(slots=True)
class AuditLogEntry:
    id: int
    mailbox_id: int | None
    mailbox_label: str = ""
    mailbox_email: str = ""
    actor: str = ""
    action: str = ""
    target_type: str = ""
    target_id: str = ""
    status: str = "success"
    details: dict[str, Any] = field(default_factory=dict)
    created_at: str = ""


@dataclass(slots=True)
class SyncJob:
    id: int
    mailbox_id: int
    method: str
    requested_by: str
    scope: dict[str, Any]
    status: str
    processed_messages: int
    cached_messages: int
    folders_synced: int
    error: str
    started_at: str
    finished_at: str


@dataclass(slots=True)
class SyncState:
    mailbox_id: int
    method: str
    folder_id: str
    last_synced_at: str
    last_message_at: str
    cached_messages: int
    status: str
    error: str
    updated_at: str


class MailboxStoreError(RuntimeError):
    """多邮箱档案存储层的统一异常。"""


class MailboxStore:
    def __init__(self, db_path: str | Path = "data/mailboxes.db") -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._fts_enabled = True
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

    def cache_folders(self, mailbox_id: int, method: str, folders: Iterable[Any]) -> list[FolderCacheEntry]:
        normalized_method = self._normalize_method_value(method)
        now = self._utc_now()
        entries = [self._normalize_folder_record(mailbox_id, normalized_method, item, now) for item in folders]

        with self._connect() as connection:
            connection.execute(
                "DELETE FROM folder_cache WHERE mailbox_id = ? AND method = ?",
                (mailbox_id, normalized_method),
            )
            if entries:
                connection.executemany(
                    """
                    INSERT INTO folder_cache (
                        mailbox_id,
                        method,
                        folder_id,
                        name,
                        display_name,
                        kind,
                        total,
                        unread,
                        is_default,
                        cached_at,
                        updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    [
                        (
                            entry.mailbox_id,
                            entry.method,
                            entry.folder_id,
                            entry.name,
                            entry.display_name,
                            entry.kind,
                            entry.total,
                            entry.unread,
                            int(entry.is_default),
                            entry.cached_at,
                            entry.updated_at,
                        )
                        for entry in entries
                    ],
                )

        return entries

    def list_cached_folders(self, mailbox_id: int, method: str | None = None) -> list[FolderCacheEntry]:
        parameters: list[Any] = [mailbox_id]
        where_clause = "WHERE mailbox_id = ?"
        if method:
            where_clause += " AND method = ?"
            parameters.append(self._normalize_method_value(method))

        with self._connect() as connection:
            rows = connection.execute(
                f"""
                SELECT
                    mailbox_id,
                    method,
                    folder_id,
                    name,
                    display_name,
                    kind,
                    total,
                    unread,
                    is_default,
                    cached_at,
                    updated_at
                FROM folder_cache
                {where_clause}
                ORDER BY method ASC, display_name COLLATE NOCASE ASC
                """,
                parameters,
            ).fetchall()
        return [self._row_to_folder_entry(row) for row in rows]

    def cache_messages(self, mailbox_id: int, method: str, messages: Iterable[Any]) -> list[CachedMessage]:
        normalized_method = self._normalize_method_value(method)
        payloads = list(messages)
        if not payloads:
            return []

        now = self._utc_now()
        cached_ids: list[str] = []

        with self._connect() as connection:
            for item in payloads:
                record = self._normalize_message_record(mailbox_id, normalized_method, item, now)
                cached_ids.append(record["provider_message_id"])
                connection.execute(
                    """
                    INSERT INTO message_cache (
                        mailbox_id,
                        method,
                        provider_message_id,
                        internet_message_id,
                        conversation_id,
                        folder_id,
                        subject,
                        sender,
                        sender_name,
                        received_at,
                        is_read,
                        is_flagged,
                        importance,
                        has_attachments,
                        preview,
                        body_text,
                        body_html,
                        to_recipients_json,
                        cc_recipients_json,
                        bcc_recipients_json,
                        headers_json,
                        in_reply_to,
                        references_json,
                        raw_json,
                        cached_at,
                        updated_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(mailbox_id, method, provider_message_id) DO UPDATE SET
                        internet_message_id = CASE
                            WHEN excluded.internet_message_id <> '' THEN excluded.internet_message_id
                            ELSE message_cache.internet_message_id
                        END,
                        conversation_id = CASE
                            WHEN excluded.conversation_id <> '' THEN excluded.conversation_id
                            ELSE message_cache.conversation_id
                        END,
                        folder_id = CASE
                            WHEN excluded.folder_id <> '' THEN excluded.folder_id
                            ELSE message_cache.folder_id
                        END,
                        subject = CASE
                            WHEN excluded.subject <> '' THEN excluded.subject
                            ELSE message_cache.subject
                        END,
                        sender = CASE
                            WHEN excluded.sender <> '' THEN excluded.sender
                            ELSE message_cache.sender
                        END,
                        sender_name = CASE
                            WHEN excluded.sender_name <> '' THEN excluded.sender_name
                            ELSE message_cache.sender_name
                        END,
                        received_at = CASE
                            WHEN excluded.received_at <> '' THEN excluded.received_at
                            ELSE message_cache.received_at
                        END,
                        is_read = excluded.is_read,
                        is_flagged = excluded.is_flagged,
                        importance = CASE
                            WHEN excluded.importance <> '' THEN excluded.importance
                            ELSE message_cache.importance
                        END,
                        has_attachments = excluded.has_attachments,
                        preview = CASE
                            WHEN excluded.preview <> '' THEN excluded.preview
                            ELSE message_cache.preview
                        END,
                        body_text = CASE
                            WHEN excluded.body_text <> '' THEN excluded.body_text
                            ELSE message_cache.body_text
                        END,
                        body_html = CASE
                            WHEN excluded.body_html <> '' THEN excluded.body_html
                            ELSE message_cache.body_html
                        END,
                        to_recipients_json = CASE
                            WHEN excluded.to_recipients_json <> '[]' THEN excluded.to_recipients_json
                            ELSE message_cache.to_recipients_json
                        END,
                        cc_recipients_json = CASE
                            WHEN excluded.cc_recipients_json <> '[]' THEN excluded.cc_recipients_json
                            ELSE message_cache.cc_recipients_json
                        END,
                        bcc_recipients_json = CASE
                            WHEN excluded.bcc_recipients_json <> '[]' THEN excluded.bcc_recipients_json
                            ELSE message_cache.bcc_recipients_json
                        END,
                        headers_json = CASE
                            WHEN excluded.headers_json <> '{}' THEN excluded.headers_json
                            ELSE message_cache.headers_json
                        END,
                        in_reply_to = CASE
                            WHEN excluded.in_reply_to <> '' THEN excluded.in_reply_to
                            ELSE message_cache.in_reply_to
                        END,
                        references_json = CASE
                            WHEN excluded.references_json <> '[]' THEN excluded.references_json
                            ELSE message_cache.references_json
                        END,
                        raw_json = CASE
                            WHEN excluded.raw_json <> '{}' THEN excluded.raw_json
                            ELSE message_cache.raw_json
                        END,
                        cached_at = excluded.cached_at,
                        updated_at = excluded.updated_at
                    """,
                    (
                        record["mailbox_id"],
                        record["method"],
                        record["provider_message_id"],
                        record["internet_message_id"],
                        record["conversation_id"],
                        record["folder_id"],
                        record["subject"],
                        record["sender"],
                        record["sender_name"],
                        record["received_at"],
                        record["is_read"],
                        record["is_flagged"],
                        record["importance"],
                        record["has_attachments"],
                        record["preview"],
                        record["body_text"],
                        record["body_html"],
                        record["to_recipients_json"],
                        record["cc_recipients_json"],
                        record["bcc_recipients_json"],
                        record["headers_json"],
                        record["in_reply_to"],
                        record["references_json"],
                        record["raw_json"],
                        record["cached_at"],
                        record["updated_at"],
                    ),
                )
                if record["attachments_provided"]:
                    self._replace_attachment_cache(
                        connection,
                        mailbox_id=mailbox_id,
                        method=normalized_method,
                        message_id=record["provider_message_id"],
                        attachments=record["attachments"],
                        now=now,
                    )
                self._refresh_search_document(connection, mailbox_id, normalized_method, record["provider_message_id"])

        return [
            item
            for message_id in cached_ids
            if (item := self.get_cached_message(mailbox_id, normalized_method, message_id)) is not None
        ]

    def cache_message(self, mailbox_id: int, method: str, message: Any) -> CachedMessage | None:
        cached = self.cache_messages(mailbox_id, method, [message])
        return cached[0] if cached else None

    def ensure_cached_message_placeholder(
        self,
        mailbox_id: int,
        method: str,
        message_id: str,
        *,
        folder_id: str = "",
    ) -> None:
        normalized_method = self._normalize_method_value(method)
        normalized_message_id = self._require_non_empty_text(message_id, "缺少邮件标识")
        now = self._utc_now()
        with self._connect() as connection:
            connection.execute(
                """
                INSERT OR IGNORE INTO message_cache (
                    mailbox_id,
                    method,
                    provider_message_id,
                    folder_id,
                    subject,
                    sender,
                    sender_name,
                    received_at,
                    is_read,
                    is_flagged,
                    importance,
                    has_attachments,
                    preview,
                    body_text,
                    body_html,
                    to_recipients_json,
                    cc_recipients_json,
                    bcc_recipients_json,
                    headers_json,
                    in_reply_to,
                    references_json,
                    raw_json,
                    cached_at,
                    updated_at
                )
                VALUES (?, ?, ?, ?, '', '', '', '', 0, 0, 'normal', 0, '', '', '', '[]', '[]', '[]', '{}', '', '[]', '{}', ?, ?)
                """,
                (mailbox_id, normalized_method, normalized_message_id, folder_id, now, now),
            )
            self._refresh_search_document(connection, mailbox_id, normalized_method, normalized_message_id)

    def get_cached_message(self, mailbox_id: int, method: str, message_id: str) -> CachedMessage | None:
        normalized_method = self._normalize_method_value(method)
        normalized_message_id = self._require_non_empty_text(message_id, "缺少邮件标识")
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT
                    c.*,
                    mb.label AS mailbox_label,
                    mb.email AS mailbox_email,
                    mm.tags_json,
                    mm.follow_up,
                    mm.notes,
                    mm.snoozed_until,
                    mm.status AS meta_status,
                    mm.updated_at AS meta_updated_at
                FROM message_cache AS c
                JOIN mailboxes AS mb
                    ON mb.id = c.mailbox_id
                LEFT JOIN message_meta AS mm
                    ON mm.mailbox_id = c.mailbox_id
                    AND mm.method = c.method
                    AND mm.provider_message_id = c.provider_message_id
                WHERE c.mailbox_id = ? AND c.method = ? AND c.provider_message_id = ?
                LIMIT 1
                """,
                (mailbox_id, normalized_method, normalized_message_id),
            ).fetchone()
            if not row:
                return None
            attachments = self._list_attachments_from_connection(
                connection,
                mailbox_id=mailbox_id,
                method=normalized_method,
                message_id=normalized_message_id,
            )
        return self._row_to_cached_message(row, attachments=attachments)

    def list_cached_messages(
        self,
        mailbox_id: int,
        *,
        method: str | None = None,
        folder: str | None = None,
        message_ids: list[str] | None = None,
        include_snoozed: bool = True,
        limit: int = 500,
    ) -> list[CachedMessage]:
        parameters: list[Any] = [mailbox_id]
        conditions = ["c.mailbox_id = ?"]
        if method:
            conditions.append("c.method = ?")
            parameters.append(self._normalize_method_value(method))
        if folder:
            conditions.append("c.folder_id = ?")
            parameters.append(folder)
        if message_ids:
            placeholders = ", ".join("?" for _ in message_ids)
            conditions.append(f"c.provider_message_id IN ({placeholders})")
            parameters.extend([self._require_non_empty_text(item, "缺少邮件标识") for item in message_ids])
        if not include_snoozed:
            conditions.append("(mm.snoozed_until IS NULL OR mm.snoozed_until = '' OR mm.snoozed_until <= ?)")
            parameters.append(self._utc_now())

        with self._connect() as connection:
            rows = connection.execute(
                f"""
                SELECT
                    c.*,
                    mb.label AS mailbox_label,
                    mb.email AS mailbox_email,
                    mm.tags_json,
                    mm.follow_up,
                    mm.notes,
                    mm.snoozed_until,
                    mm.status AS meta_status,
                    mm.updated_at AS meta_updated_at
                FROM message_cache AS c
                JOIN mailboxes AS mb
                    ON mb.id = c.mailbox_id
                LEFT JOIN message_meta AS mm
                    ON mm.mailbox_id = c.mailbox_id
                    AND mm.method = c.method
                    AND mm.provider_message_id = c.provider_message_id
                WHERE {' AND '.join(conditions)}
                ORDER BY c.received_at DESC, c.id DESC
                LIMIT ?
                """,
                [*parameters, max(limit, 1)],
            ).fetchall()
            attachments_map = self._list_attachments_map_from_connection(
                connection,
                mailbox_id=mailbox_id,
                method=self._normalize_method_value(method) if method else None,
                message_ids=[str(row["provider_message_id"]) for row in rows],
            )

        return [
            self._row_to_cached_message(
                row,
                attachments=attachments_map.get(str(row["provider_message_id"]), []),
            )
            for row in rows
        ]

    def update_cached_message_state(
        self,
        mailbox_id: int,
        method: str,
        message_id: str,
        *,
        is_read: bool | None = None,
        is_flagged: bool | None = None,
        folder_id: str | None = None,
        importance: str | None = None,
    ) -> CachedMessage | None:
        normalized_method = self._normalize_method_value(method)
        normalized_message_id = self._require_non_empty_text(message_id, "缺少邮件标识")
        assignments: list[str] = []
        parameters: list[Any] = []

        if is_read is not None:
            assignments.append("is_read = ?")
            parameters.append(int(is_read))
        if is_flagged is not None:
            assignments.append("is_flagged = ?")
            parameters.append(int(is_flagged))
        if folder_id is not None:
            assignments.append("folder_id = ?")
            parameters.append(folder_id)
        if importance is not None:
            assignments.append("importance = ?")
            parameters.append(importance or "normal")

        if not assignments:
            return self.get_cached_message(mailbox_id, normalized_method, normalized_message_id)

        now = self._utc_now()
        with self._connect() as connection:
            connection.execute(
                f"""
                UPDATE message_cache
                SET {', '.join(assignments)}, updated_at = ?
                WHERE mailbox_id = ? AND method = ? AND provider_message_id = ?
                """,
                [*parameters, now, mailbox_id, normalized_method, normalized_message_id],
            )
            self._refresh_search_document(connection, mailbox_id, normalized_method, normalized_message_id)
        return self.get_cached_message(mailbox_id, normalized_method, normalized_message_id)

    def remove_cached_message(self, mailbox_id: int, method: str, message_id: str) -> bool:
        normalized_method = self._normalize_method_value(method)
        normalized_message_id = self._require_non_empty_text(message_id, "缺少邮件标识")
        document_id = self._build_document_id(mailbox_id, normalized_method, normalized_message_id)
        with self._connect() as connection:
            connection.execute(
                "DELETE FROM attachment_cache WHERE mailbox_id = ? AND method = ? AND provider_message_id = ?",
                (mailbox_id, normalized_method, normalized_message_id),
            )
            connection.execute(
                "DELETE FROM message_meta WHERE mailbox_id = ? AND method = ? AND provider_message_id = ?",
                (mailbox_id, normalized_method, normalized_message_id),
            )
            cursor = connection.execute(
                "DELETE FROM message_cache WHERE mailbox_id = ? AND method = ? AND provider_message_id = ?",
                (mailbox_id, normalized_method, normalized_message_id),
            )
            connection.execute("DELETE FROM message_search WHERE document_id = ?", (document_id,))
        return cursor.rowcount > 0

    def list_message_meta_map(self, mailbox_id: int, method: str, message_ids: list[str]) -> dict[str, MessageMeta]:
        normalized_method = self._normalize_method_value(method)
        unique_ids = [self._require_non_empty_text(item, "缺少邮件标识") for item in dict.fromkeys(message_ids)]
        if not unique_ids:
            return {}
        placeholders = ", ".join("?" for _ in unique_ids)
        with self._connect() as connection:
            rows = connection.execute(
                f"""
                SELECT
                    mailbox_id,
                    method,
                    provider_message_id,
                    tags_json,
                    follow_up,
                    notes,
                    snoozed_until,
                    status,
                    updated_at
                FROM message_meta
                WHERE mailbox_id = ? AND method = ? AND provider_message_id IN ({placeholders})
                """,
                [mailbox_id, normalized_method, *unique_ids],
            ).fetchall()
        return {
            str(row["provider_message_id"]): self._row_to_message_meta(row)
            for row in rows
        }

    def get_message_meta(self, mailbox_id: int, method: str, message_id: str) -> MessageMeta | None:
        normalized_method = self._normalize_method_value(method)
        normalized_message_id = self._require_non_empty_text(message_id, "缺少邮件标识")
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT
                    mailbox_id,
                    method,
                    provider_message_id,
                    tags_json,
                    follow_up,
                    notes,
                    snoozed_until,
                    status,
                    updated_at
                FROM message_meta
                WHERE mailbox_id = ? AND method = ? AND provider_message_id = ?
                LIMIT 1
                """,
                (mailbox_id, normalized_method, normalized_message_id),
            ).fetchone()
        return self._row_to_message_meta(row) if row else None

    def update_message_meta(
        self,
        mailbox_id: int,
        method: str,
        message_id: str,
        *,
        tags: list[str] | None = None,
        follow_up: str | None = None,
        notes: str | None = None,
        snoozed_until: str | None = None,
        status: str | None = None,
    ) -> MessageMeta:
        normalized_method = self._normalize_method_value(method)
        normalized_message_id = self._require_non_empty_text(message_id, "缺少邮件标识")
        current = self.get_message_meta(mailbox_id, normalized_method, normalized_message_id)
        now = self._utc_now()

        next_tags = self._normalize_tags(tags) if tags is not None else (current.tags if current else [])
        next_follow_up = self._normalize_optional_string(follow_up) if follow_up is not None else (current.follow_up if current else "")
        next_notes = self._normalize_optional_string(notes) if notes is not None else (current.notes if current else "")
        next_snoozed_until = (
            self._normalize_optional_string(snoozed_until)
            if snoozed_until is not None
            else (current.snoozed_until if current else "")
        )
        next_status = self._normalize_optional_string(status) if status is not None else (current.status if current else "active")
        if next_snoozed_until and next_status == "active":
            next_status = "snoozed"

        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO message_meta (
                    mailbox_id,
                    method,
                    provider_message_id,
                    tags_json,
                    tags_text,
                    follow_up,
                    notes,
                    snoozed_until,
                    status,
                    updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(mailbox_id, method, provider_message_id) DO UPDATE SET
                    tags_json = excluded.tags_json,
                    tags_text = excluded.tags_text,
                    follow_up = excluded.follow_up,
                    notes = excluded.notes,
                    snoozed_until = excluded.snoozed_until,
                    status = excluded.status,
                    updated_at = excluded.updated_at
                """,
                (
                    mailbox_id,
                    normalized_method,
                    normalized_message_id,
                    self._json_dumps(next_tags),
                    " ".join(next_tags),
                    next_follow_up,
                    next_notes,
                    next_snoozed_until,
                    next_status or "active",
                    now,
                ),
            )
            self._refresh_search_document(connection, mailbox_id, normalized_method, normalized_message_id)

        return (
            self.get_message_meta(mailbox_id, normalized_method, normalized_message_id)
            or MessageMeta(
                mailbox_id=mailbox_id,
                method=normalized_method,
                message_id=normalized_message_id,
                tags=next_tags,
                follow_up=next_follow_up,
                notes=next_notes,
                snoozed_until=next_snoozed_until,
                status=next_status or "active",
                updated_at=now,
            )
        )

    def replace_attachment_cache(
        self,
        mailbox_id: int,
        method: str,
        message_id: str,
        attachments: Iterable[Any],
    ) -> list[CachedAttachment]:
        normalized_method = self._normalize_method_value(method)
        normalized_message_id = self._require_non_empty_text(message_id, "缺少邮件标识")
        attachment_list = list(attachments)
        now = self._utc_now()

        with self._connect() as connection:
            self._replace_attachment_cache(
                connection,
                mailbox_id=mailbox_id,
                method=normalized_method,
                message_id=normalized_message_id,
                attachments=attachment_list,
                now=now,
            )

        return self.list_cached_attachments(mailbox_id, normalized_method, normalized_message_id)

    def upsert_attachment_content(
        self,
        mailbox_id: int,
        method: str,
        message_id: str,
        attachment: Any,
    ) -> CachedAttachment | None:
        normalized_method = self._normalize_method_value(method)
        normalized_message_id = self._require_non_empty_text(message_id, "缺少邮件标识")
        now = self._utc_now()
        record = self._normalize_attachment_record(mailbox_id, normalized_method, normalized_message_id, attachment, now)
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO attachment_cache (
                    mailbox_id,
                    method,
                    provider_message_id,
                    attachment_id,
                    name,
                    content_type,
                    size,
                    is_inline,
                    content_base64,
                    cached_at,
                    updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(mailbox_id, method, provider_message_id, attachment_id) DO UPDATE SET
                    name = excluded.name,
                    content_type = excluded.content_type,
                    size = excluded.size,
                    is_inline = excluded.is_inline,
                    content_base64 = CASE
                        WHEN excluded.content_base64 <> '' THEN excluded.content_base64
                        ELSE attachment_cache.content_base64
                    END,
                    cached_at = excluded.cached_at,
                    updated_at = excluded.updated_at
                """,
                (
                    record.mailbox_id,
                    record.method,
                    record.message_id,
                    record.attachment_id,
                    record.name,
                    record.content_type,
                    record.size,
                    int(record.is_inline),
                    record.content_base64,
                    record.cached_at,
                    record.updated_at,
                ),
            )
        return self.get_cached_attachment(mailbox_id, normalized_method, normalized_message_id, record.attachment_id)

    def list_cached_attachments(self, mailbox_id: int, method: str, message_id: str) -> list[CachedAttachment]:
        normalized_method = self._normalize_method_value(method)
        normalized_message_id = self._require_non_empty_text(message_id, "缺少邮件标识")
        with self._connect() as connection:
            return self._list_attachments_from_connection(
                connection,
                mailbox_id=mailbox_id,
                method=normalized_method,
                message_id=normalized_message_id,
            )

    def get_cached_attachment(
        self,
        mailbox_id: int,
        method: str,
        message_id: str,
        attachment_id: str,
    ) -> CachedAttachment | None:
        normalized_method = self._normalize_method_value(method)
        normalized_message_id = self._require_non_empty_text(message_id, "缺少邮件标识")
        normalized_attachment_id = self._require_non_empty_text(attachment_id, "缺少附件标识")
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT
                    mailbox_id,
                    method,
                    provider_message_id,
                    attachment_id,
                    name,
                    content_type,
                    size,
                    is_inline,
                    content_base64,
                    cached_at,
                    updated_at
                FROM attachment_cache
                WHERE mailbox_id = ? AND method = ? AND provider_message_id = ? AND attachment_id = ?
                LIMIT 1
                """,
                (mailbox_id, normalized_method, normalized_message_id, normalized_attachment_id),
            ).fetchone()
        return self._row_to_attachment(row) if row else None

    def search_messages(
        self,
        query: str | None,
        *,
        mailbox_ids: list[int] | None = None,
        method: str | None = None,
        folder: str | None = None,
        tag: str | None = None,
        unread_only: bool = False,
        flagged_only: bool = False,
        has_attachments_only: bool = False,
        include_snoozed: bool = True,
        page: int,
        page_size: int,
        sort_order: str = "desc",
    ) -> tuple[list[MessageSearchResult], int]:
        normalized_query = self._normalize_optional_string(query)
        joins = [
            "JOIN mailboxes AS mb ON mb.id = c.mailbox_id",
            """
            LEFT JOIN message_meta AS mm
                ON mm.mailbox_id = c.mailbox_id
                AND mm.method = c.method
                AND mm.provider_message_id = c.provider_message_id
            """,
        ]
        conditions = ["1 = 1"]
        parameters: list[Any] = []

        if mailbox_ids:
            placeholders = ", ".join("?" for _ in mailbox_ids)
            conditions.append(f"c.mailbox_id IN ({placeholders})")
            parameters.extend(mailbox_ids)
        if method:
            conditions.append("c.method = ?")
            parameters.append(self._normalize_method_value(method))
        if folder:
            conditions.append("c.folder_id = ?")
            parameters.append(folder)
        if tag:
            conditions.append("lower(COALESCE(mm.tags_text, '')) LIKE ?")
            parameters.append(f"%{tag.strip().lower()}%")
        if unread_only:
            conditions.append("c.is_read = 0")
        if flagged_only:
            conditions.append("c.is_flagged = 1")
        if has_attachments_only:
            conditions.append("c.has_attachments = 1")
        if not include_snoozed:
            conditions.append("(mm.snoozed_until IS NULL OR mm.snoozed_until = '' OR mm.snoozed_until <= ?)")
            parameters.append(self._utc_now())

        if normalized_query:
            if self._fts_enabled:
                joins.append(
                    """
                    JOIN message_search
                        ON message_search.document_id = (
                            CAST(c.mailbox_id AS TEXT) || '|' || c.method || '|' || c.provider_message_id
                        )
                    """
                )
                conditions.append("message_search MATCH ?")
                parameters.append(self._build_fts_query(normalized_query))
            else:
                like_pattern = f"%{normalized_query.lower()}%"
                conditions.append(
                    """
                    (
                        lower(c.subject) LIKE ?
                        OR lower(c.sender) LIKE ?
                        OR lower(c.preview) LIKE ?
                        OR lower(c.body_text) LIKE ?
                        OR lower(COALESCE(mm.notes, '')) LIKE ?
                        OR lower(COALESCE(mm.tags_text, '')) LIKE ?
                    )
                    """
                )
                parameters.extend([like_pattern] * 6)

        where_clause = " AND ".join(conditions)
        order_direction = "ASC" if str(sort_order).strip().casefold() == "asc" else "DESC"
        offset = max(page - 1, 0) * page_size
        join_clause = "\n".join(joins)

        with self._connect() as connection:
            total = int(
                connection.execute(
                    f"""
                    SELECT COUNT(*) AS total
                    FROM message_cache AS c
                    {join_clause}
                    WHERE {where_clause}
                    """,
                    parameters,
                ).fetchone()["total"]
            )
            rows = connection.execute(
                f"""
                SELECT
                    c.*,
                    mb.label AS mailbox_label,
                    mb.email AS mailbox_email,
                    mm.tags_json,
                    mm.follow_up,
                    mm.notes,
                    mm.snoozed_until,
                    mm.status AS meta_status,
                    mm.updated_at AS meta_updated_at
                FROM message_cache AS c
                {join_clause}
                WHERE {where_clause}
                ORDER BY c.received_at {order_direction}, c.id DESC
                LIMIT ? OFFSET ?
                """,
                [*parameters, page_size, offset],
            ).fetchall()

        return [self._row_to_search_result(row) for row in rows], total

    def list_thread_messages(
        self,
        mailbox_id: int,
        *,
        method: str | None = None,
        conversation_id: str | None = None,
        message_id: str | None = None,
    ) -> list[CachedMessage]:
        anchor_method = self._normalize_method_value(method) if method else None
        anchor_message_id = self._normalize_optional_string(message_id)
        target_conversation_id = self._normalize_optional_string(conversation_id)
        target_internet_message_id = ""

        if not target_conversation_id and anchor_message_id:
            candidates = self.list_cached_messages(
                mailbox_id,
                method=anchor_method,
                message_ids=[anchor_message_id],
                limit=1,
            )
            if candidates:
                target_conversation_id = candidates[0].conversation_id or ""
                target_internet_message_id = candidates[0].internet_message_id or ""

        conditions = ["c.mailbox_id = ?"]
        parameters: list[Any] = [mailbox_id]
        if anchor_method:
            conditions.append("c.method = ?")
            parameters.append(anchor_method)

        if target_conversation_id:
            conditions.append("c.conversation_id = ?")
            parameters.append(target_conversation_id)
        elif target_internet_message_id:
            conditions.append(
                """
                (
                    c.internet_message_id = ?
                    OR c.in_reply_to = ?
                    OR lower(c.references_json) LIKE ?
                )
                """
            )
            parameters.extend(
                [
                    target_internet_message_id,
                    target_internet_message_id,
                    f"%{target_internet_message_id.lower()}%",
                ]
            )
        elif anchor_message_id:
            conditions.append("c.provider_message_id = ?")
            parameters.append(anchor_message_id)
        else:
            return []

        with self._connect() as connection:
            rows = connection.execute(
                f"""
                SELECT
                    c.*,
                    mb.label AS mailbox_label,
                    mb.email AS mailbox_email,
                    mm.tags_json,
                    mm.follow_up,
                    mm.notes,
                    mm.snoozed_until,
                    mm.status AS meta_status,
                    mm.updated_at AS meta_updated_at
                FROM message_cache AS c
                JOIN mailboxes AS mb
                    ON mb.id = c.mailbox_id
                LEFT JOIN message_meta AS mm
                    ON mm.mailbox_id = c.mailbox_id
                    AND mm.method = c.method
                    AND mm.provider_message_id = c.provider_message_id
                WHERE {' AND '.join(conditions)}
                ORDER BY c.received_at ASC, c.id ASC
                """,
                parameters,
            ).fetchall()
            attachments_map = self._list_attachments_map_from_connection(
                connection,
                mailbox_id=mailbox_id,
                method=anchor_method,
                message_ids=[str(row["provider_message_id"]) for row in rows],
            )

        return [
            self._row_to_cached_message(
                row,
                attachments=attachments_map.get(str(row["provider_message_id"]), []),
            )
            for row in rows
        ]

    def create_rule(
        self,
        mailbox_id: int,
        *,
        name: str,
        enabled: bool,
        priority: int,
        conditions: dict[str, Any],
        actions: dict[str, Any],
    ) -> SavedRule:
        normalized_name = self._require_non_empty_text(name, "规则名称不能为空")
        now = self._utc_now()
        with self._connect() as connection:
            cursor = connection.execute(
                """
                INSERT INTO saved_rules (
                    mailbox_id,
                    name,
                    enabled,
                    priority,
                    conditions_json,
                    actions_json,
                    created_at,
                    updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    mailbox_id,
                    normalized_name,
                    int(enabled),
                    int(priority),
                    self._json_dumps(conditions or {}),
                    self._json_dumps(actions or {}),
                    now,
                    now,
                ),
            )
            rule_id = int(cursor.lastrowid)
        rule = self.get_rule(rule_id, mailbox_id)
        if not rule:
            raise MailboxStoreError("规则创建后读取失败")
        return rule

    def get_rule(self, rule_id: int, mailbox_id: int | None = None) -> SavedRule | None:
        parameters: list[Any] = [rule_id]
        where_clause = "WHERE id = ?"
        if mailbox_id is not None:
            where_clause += " AND mailbox_id = ?"
            parameters.append(mailbox_id)
        with self._connect() as connection:
            row = connection.execute(
                f"""
                SELECT
                    id,
                    mailbox_id,
                    name,
                    enabled,
                    priority,
                    conditions_json,
                    actions_json,
                    created_at,
                    updated_at
                FROM saved_rules
                {where_clause}
                LIMIT 1
                """,
                parameters,
            ).fetchone()
        return self._row_to_rule(row) if row else None

    def list_rules(self, mailbox_id: int, *, enabled_only: bool = False) -> list[SavedRule]:
        parameters: list[Any] = [mailbox_id]
        where_clause = "WHERE mailbox_id = ?"
        if enabled_only:
            where_clause += " AND enabled = 1"
        with self._connect() as connection:
            rows = connection.execute(
                f"""
                SELECT
                    id,
                    mailbox_id,
                    name,
                    enabled,
                    priority,
                    conditions_json,
                    actions_json,
                    created_at,
                    updated_at
                FROM saved_rules
                {where_clause}
                ORDER BY priority ASC, id ASC
                """,
                parameters,
            ).fetchall()
        return [self._row_to_rule(row) for row in rows]

    def update_rule(
        self,
        rule_id: int,
        mailbox_id: int,
        payload: dict[str, Any],
    ) -> SavedRule:
        existing = self.get_rule(rule_id, mailbox_id)
        if not existing:
            raise MailboxStoreError("规则不存在")

        now = self._utc_now()
        name = self._normalize_optional_string(payload.get("name")) or existing.name
        enabled = bool(payload.get("enabled")) if "enabled" in payload else existing.enabled
        priority = int(payload.get("priority")) if "priority" in payload else existing.priority
        conditions = payload.get("conditions") if "conditions" in payload else existing.conditions
        actions = payload.get("actions") if "actions" in payload else existing.actions
        if not isinstance(conditions, dict):
            raise MailboxStoreError("conditions 必须是对象")
        if not isinstance(actions, dict):
            raise MailboxStoreError("actions 必须是对象")

        with self._connect() as connection:
            connection.execute(
                """
                UPDATE saved_rules
                SET
                    name = ?,
                    enabled = ?,
                    priority = ?,
                    conditions_json = ?,
                    actions_json = ?,
                    updated_at = ?
                WHERE id = ? AND mailbox_id = ?
                """,
                (
                    name,
                    int(enabled),
                    priority,
                    self._json_dumps(conditions),
                    self._json_dumps(actions),
                    now,
                    rule_id,
                    mailbox_id,
                ),
            )
        updated = self.get_rule(rule_id, mailbox_id)
        if not updated:
            raise MailboxStoreError("规则更新后读取失败")
        return updated

    def delete_rule(self, rule_id: int, mailbox_id: int) -> bool:
        with self._connect() as connection:
            cursor = connection.execute(
                "DELETE FROM saved_rules WHERE id = ? AND mailbox_id = ?",
                (rule_id, mailbox_id),
            )
        return cursor.rowcount > 0

    def record_audit_log(
        self,
        *,
        mailbox_id: int | None,
        actor: str,
        action: str,
        target_type: str,
        target_id: str = "",
        status: str = "success",
        details: dict[str, Any] | None = None,
    ) -> AuditLogEntry:
        now = self._utc_now()
        with self._connect() as connection:
            cursor = connection.execute(
                """
                INSERT INTO audit_logs (
                    mailbox_id,
                    actor,
                    action,
                    target_type,
                    target_id,
                    status,
                    details_json,
                    created_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    mailbox_id,
                    self._normalize_optional_string(actor) or "system",
                    self._require_non_empty_text(action, "审计动作不能为空"),
                    self._require_non_empty_text(target_type, "审计目标类型不能为空"),
                    self._normalize_optional_string(target_id) or "",
                    self._normalize_optional_string(status) or "success",
                    self._json_dumps(details or {}),
                    now,
                ),
            )
            log_id = int(cursor.lastrowid)
        log = self.get_audit_log(log_id)
        if not log:
            raise MailboxStoreError("审计日志写入后读取失败")
        return log

    def get_audit_log(self, log_id: int) -> AuditLogEntry | None:
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT
                    a.id,
                    a.mailbox_id,
                    COALESCE(m.label, '') AS mailbox_label,
                    COALESCE(m.email, '') AS mailbox_email,
                    a.actor,
                    a.action,
                    a.target_type,
                    a.target_id,
                    a.status,
                    a.details_json,
                    a.created_at
                FROM audit_logs AS a
                LEFT JOIN mailboxes AS m
                    ON m.id = a.mailbox_id
                WHERE a.id = ?
                LIMIT 1
                """,
                (log_id,),
            ).fetchone()
        return self._row_to_audit_log(row) if row else None

    def list_audit_logs(
        self,
        *,
        mailbox_id: int | None = None,
        action: str | None = None,
        page: int,
        page_size: int,
    ) -> tuple[list[AuditLogEntry], int]:
        conditions = ["1 = 1"]
        parameters: list[Any] = []
        if mailbox_id is not None:
            conditions.append("a.mailbox_id = ?")
            parameters.append(mailbox_id)
        if action:
            conditions.append("a.action = ?")
            parameters.append(action.strip())

        where_clause = " AND ".join(conditions)
        offset = max(page - 1, 0) * page_size

        with self._connect() as connection:
            total = int(
                connection.execute(
                    f"SELECT COUNT(*) AS total FROM audit_logs AS a WHERE {where_clause}",
                    parameters,
                ).fetchone()["total"]
            )
            rows = connection.execute(
                f"""
                SELECT
                    a.id,
                    a.mailbox_id,
                    COALESCE(m.label, '') AS mailbox_label,
                    COALESCE(m.email, '') AS mailbox_email,
                    a.actor,
                    a.action,
                    a.target_type,
                    a.target_id,
                    a.status,
                    a.details_json,
                    a.created_at
                FROM audit_logs AS a
                LEFT JOIN mailboxes AS m
                    ON m.id = a.mailbox_id
                WHERE {where_clause}
                ORDER BY a.id DESC
                LIMIT ? OFFSET ?
                """,
                [*parameters, page_size, offset],
            ).fetchall()
        return [self._row_to_audit_log(row) for row in rows], total

    def create_sync_job(
        self,
        *,
        mailbox_id: int,
        method: str,
        requested_by: str,
        scope: dict[str, Any] | None = None,
    ) -> SyncJob:
        normalized_method = self._normalize_method_value(method)
        now = self._utc_now()
        with self._connect() as connection:
            cursor = connection.execute(
                """
                INSERT INTO sync_jobs (
                    mailbox_id,
                    method,
                    requested_by,
                    scope_json,
                    status,
                    processed_messages,
                    cached_messages,
                    folders_synced,
                    error,
                    started_at,
                    finished_at
                )
                VALUES (?, ?, ?, ?, 'running', 0, 0, 0, '', ?, '')
                """,
                (
                    mailbox_id,
                    normalized_method,
                    self._normalize_optional_string(requested_by) or "system",
                    self._json_dumps(scope or {}),
                    now,
                ),
            )
            job_id = int(cursor.lastrowid)
        job = self.get_sync_job(job_id)
        if not job:
            raise MailboxStoreError("同步任务创建后读取失败")
        return job

    def get_sync_job(self, job_id: int) -> SyncJob | None:
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT
                    id,
                    mailbox_id,
                    method,
                    requested_by,
                    scope_json,
                    status,
                    processed_messages,
                    cached_messages,
                    folders_synced,
                    error,
                    started_at,
                    finished_at
                FROM sync_jobs
                WHERE id = ?
                LIMIT 1
                """,
                (job_id,),
            ).fetchone()
        return self._row_to_sync_job(row) if row else None

    def update_sync_job(
        self,
        job_id: int,
        *,
        status: str,
        processed_messages: int,
        cached_messages: int,
        folders_synced: int,
        error: str = "",
    ) -> SyncJob:
        now = self._utc_now()
        with self._connect() as connection:
            connection.execute(
                """
                UPDATE sync_jobs
                SET
                    status = ?,
                    processed_messages = ?,
                    cached_messages = ?,
                    folders_synced = ?,
                    error = ?,
                    finished_at = ?
                WHERE id = ?
                """,
                (
                    self._normalize_optional_string(status) or "completed",
                    max(processed_messages, 0),
                    max(cached_messages, 0),
                    max(folders_synced, 0),
                    self._normalize_optional_string(error) or "",
                    now,
                    job_id,
                ),
            )
        job = self.get_sync_job(job_id)
        if not job:
            raise MailboxStoreError("同步任务更新后读取失败")
        return job

    def list_sync_jobs(self, mailbox_id: int, *, limit: int = 20) -> list[SyncJob]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT
                    id,
                    mailbox_id,
                    method,
                    requested_by,
                    scope_json,
                    status,
                    processed_messages,
                    cached_messages,
                    folders_synced,
                    error,
                    started_at,
                    finished_at
                FROM sync_jobs
                WHERE mailbox_id = ?
                ORDER BY id DESC
                LIMIT ?
                """,
                (mailbox_id, max(limit, 1)),
            ).fetchall()
        return [self._row_to_sync_job(row) for row in rows]

    def upsert_sync_state(
        self,
        *,
        mailbox_id: int,
        method: str,
        folder_id: str,
        cached_messages: int,
        last_message_at: str,
        status: str,
        error: str = "",
    ) -> SyncState:
        normalized_method = self._normalize_method_value(method)
        normalized_folder_id = self._normalize_optional_string(folder_id) or "__mailbox__"
        now = self._utc_now()
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO sync_state (
                    mailbox_id,
                    method,
                    folder_id,
                    last_synced_at,
                    last_message_at,
                    cached_messages,
                    status,
                    error,
                    updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(mailbox_id, method, folder_id) DO UPDATE SET
                    last_synced_at = excluded.last_synced_at,
                    last_message_at = excluded.last_message_at,
                    cached_messages = excluded.cached_messages,
                    status = excluded.status,
                    error = excluded.error,
                    updated_at = excluded.updated_at
                """,
                (
                    mailbox_id,
                    normalized_method,
                    normalized_folder_id,
                    now,
                    self._normalize_optional_string(last_message_at) or "",
                    max(cached_messages, 0),
                    self._normalize_optional_string(status) or "completed",
                    self._normalize_optional_string(error) or "",
                    now,
                ),
            )
        state = self.get_sync_state(mailbox_id, normalized_method, normalized_folder_id)
        if not state:
            raise MailboxStoreError("同步状态写入后读取失败")
        return state

    def get_sync_state(self, mailbox_id: int, method: str, folder_id: str) -> SyncState | None:
        normalized_method = self._normalize_method_value(method)
        normalized_folder_id = self._normalize_optional_string(folder_id) or "__mailbox__"
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT
                    mailbox_id,
                    method,
                    folder_id,
                    last_synced_at,
                    last_message_at,
                    cached_messages,
                    status,
                    error,
                    updated_at
                FROM sync_state
                WHERE mailbox_id = ? AND method = ? AND folder_id = ?
                LIMIT 1
                """,
                (mailbox_id, normalized_method, normalized_folder_id),
            ).fetchone()
        return self._row_to_sync_state(row) if row else None

    def list_sync_states(self, mailbox_id: int, method: str | None = None) -> list[SyncState]:
        parameters: list[Any] = [mailbox_id]
        where_clause = "WHERE mailbox_id = ?"
        if method:
            where_clause += " AND method = ?"
            parameters.append(self._normalize_method_value(method))

        with self._connect() as connection:
            rows = connection.execute(
                f"""
                SELECT
                    mailbox_id,
                    method,
                    folder_id,
                    last_synced_at,
                    last_message_at,
                    cached_messages,
                    status,
                    error,
                    updated_at
                FROM sync_state
                {where_clause}
                ORDER BY method ASC, folder_id ASC
                """,
                parameters,
            ).fetchall()
        return [self._row_to_sync_state(row) for row in rows]

    def create_compose_attachment(
        self,
        mailbox_id: int,
        *,
        method: str,
        file_name: str,
        content_type: str,
        content_base64: str,
    ) -> dict[str, Any]:
        normalized_method = self._normalize_method_value(method)
        normalized_name = self._require_non_empty_text(file_name, "附件文件名不能为空")
        normalized_content = self._require_non_empty_text(content_base64, "附件内容不能为空")
        now = self._utc_now()
        token = uuid.uuid4().hex

        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO compose_attachments (
                    token,
                    mailbox_id,
                    method,
                    file_name,
                    content_type,
                    content_base64,
                    created_at,
                    updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    token,
                    mailbox_id,
                    normalized_method,
                    normalized_name,
                    self._normalize_optional_string(content_type) or "application/octet-stream",
                    normalized_content,
                    now,
                    now,
                ),
            )
        return {
            "token": token,
            "mailbox_id": mailbox_id,
            "method": normalized_method,
            "file_name": normalized_name,
            "content_type": self._normalize_optional_string(content_type) or "application/octet-stream",
            "size": len(normalized_content),
            "created_at": now,
        }

    def get_compose_attachments(
        self,
        mailbox_id: int,
        *,
        method: str,
        tokens: list[str],
    ) -> list[dict[str, Any]]:
        normalized_method = self._normalize_method_value(method)
        unique_tokens = [self._require_non_empty_text(item, "附件令牌不能为空") for item in dict.fromkeys(tokens)]
        if not unique_tokens:
            return []
        placeholders = ", ".join("?" for _ in unique_tokens)
        with self._connect() as connection:
            rows = connection.execute(
                f"""
                SELECT
                    token,
                    mailbox_id,
                    method,
                    file_name,
                    content_type,
                    content_base64,
                    created_at,
                    updated_at
                FROM compose_attachments
                WHERE mailbox_id = ? AND method = ? AND token IN ({placeholders})
                ORDER BY created_at ASC, token ASC
                """,
                [mailbox_id, normalized_method, *unique_tokens],
            ).fetchall()
        return [
            {
                "token": str(row["token"]),
                "mailbox_id": int(row["mailbox_id"]),
                "method": str(row["method"]),
                "name": str(row["file_name"]),
                "content_type": str(row["content_type"] or "application/octet-stream"),
                "content_base64": str(row["content_base64"] or ""),
                "created_at": str(row["created_at"]),
                "updated_at": str(row["updated_at"]),
            }
            for row in rows
        ]

    def delete_compose_attachment(self, token: str) -> bool:
        normalized_token = self._require_non_empty_text(token, "附件令牌不能为空")
        with self._connect() as connection:
            cursor = connection.execute("DELETE FROM compose_attachments WHERE token = ?", (normalized_token,))
        return cursor.rowcount > 0

    def get_cached_attachment_content(
        self,
        mailbox_id: int,
        method: str,
        message_id: str,
        attachment_id: str,
    ) -> dict[str, Any] | None:
        cached = self.get_cached_attachment(mailbox_id, method, message_id, attachment_id)
        if not cached or not cached.content_base64:
            return None
        return {
            "name": cached.name,
            "content_type": cached.content_type,
            "content_base64": cached.content_base64,
            "size": cached.size,
        }

    def _initialize_schema(self) -> None:
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
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS folder_cache (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    folder_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    display_name TEXT NOT NULL,
                    kind TEXT NOT NULL DEFAULT 'custom',
                    total INTEGER NOT NULL DEFAULT 0,
                    unread INTEGER NOT NULL DEFAULT 0,
                    is_default INTEGER NOT NULL DEFAULT 0,
                    cached_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(mailbox_id, method, folder_id),
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS message_cache (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    provider_message_id TEXT NOT NULL,
                    internet_message_id TEXT NOT NULL DEFAULT '',
                    conversation_id TEXT NOT NULL DEFAULT '',
                    folder_id TEXT NOT NULL DEFAULT '',
                    subject TEXT NOT NULL DEFAULT '',
                    sender TEXT NOT NULL DEFAULT '',
                    sender_name TEXT NOT NULL DEFAULT '',
                    received_at TEXT NOT NULL DEFAULT '',
                    is_read INTEGER NOT NULL DEFAULT 0,
                    is_flagged INTEGER NOT NULL DEFAULT 0,
                    importance TEXT NOT NULL DEFAULT 'normal',
                    has_attachments INTEGER NOT NULL DEFAULT 0,
                    preview TEXT NOT NULL DEFAULT '',
                    body_text TEXT NOT NULL DEFAULT '',
                    body_html TEXT DEFAULT '',
                    to_recipients_json TEXT NOT NULL DEFAULT '[]',
                    cc_recipients_json TEXT NOT NULL DEFAULT '[]',
                    bcc_recipients_json TEXT NOT NULL DEFAULT '[]',
                    headers_json TEXT NOT NULL DEFAULT '{}',
                    in_reply_to TEXT NOT NULL DEFAULT '',
                    references_json TEXT NOT NULL DEFAULT '[]',
                    raw_json TEXT NOT NULL DEFAULT '{}',
                    cached_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(mailbox_id, method, provider_message_id),
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS attachment_cache (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    provider_message_id TEXT NOT NULL,
                    attachment_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    content_type TEXT NOT NULL,
                    size INTEGER NOT NULL DEFAULT 0,
                    is_inline INTEGER NOT NULL DEFAULT 0,
                    content_base64 TEXT NOT NULL DEFAULT '',
                    cached_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(mailbox_id, method, provider_message_id, attachment_id),
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS message_meta (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    provider_message_id TEXT NOT NULL,
                    tags_json TEXT NOT NULL DEFAULT '[]',
                    tags_text TEXT NOT NULL DEFAULT '',
                    follow_up TEXT NOT NULL DEFAULT '',
                    notes TEXT NOT NULL DEFAULT '',
                    snoozed_until TEXT NOT NULL DEFAULT '',
                    status TEXT NOT NULL DEFAULT 'active',
                    updated_at TEXT NOT NULL,
                    UNIQUE(mailbox_id, method, provider_message_id),
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS saved_rules (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    enabled INTEGER NOT NULL DEFAULT 1,
                    priority INTEGER NOT NULL DEFAULT 100,
                    conditions_json TEXT NOT NULL DEFAULT '{}',
                    actions_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS audit_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER,
                    actor TEXT NOT NULL DEFAULT 'system',
                    action TEXT NOT NULL,
                    target_type TEXT NOT NULL,
                    target_id TEXT NOT NULL DEFAULT '',
                    status TEXT NOT NULL DEFAULT 'success',
                    details_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL,
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS sync_jobs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    requested_by TEXT NOT NULL DEFAULT 'system',
                    scope_json TEXT NOT NULL DEFAULT '{}',
                    status TEXT NOT NULL DEFAULT 'running',
                    processed_messages INTEGER NOT NULL DEFAULT 0,
                    cached_messages INTEGER NOT NULL DEFAULT 0,
                    folders_synced INTEGER NOT NULL DEFAULT 0,
                    error TEXT NOT NULL DEFAULT '',
                    started_at TEXT NOT NULL,
                    finished_at TEXT NOT NULL DEFAULT '',
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS sync_state (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    folder_id TEXT NOT NULL,
                    last_synced_at TEXT NOT NULL DEFAULT '',
                    last_message_at TEXT NOT NULL DEFAULT '',
                    cached_messages INTEGER NOT NULL DEFAULT 0,
                    status TEXT NOT NULL DEFAULT 'idle',
                    error TEXT NOT NULL DEFAULT '',
                    updated_at TEXT NOT NULL,
                    UNIQUE(mailbox_id, method, folder_id),
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS compose_attachments (
                    token TEXT PRIMARY KEY,
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    file_name TEXT NOT NULL,
                    content_type TEXT NOT NULL,
                    content_base64 TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_message_cache_lookup ON message_cache(mailbox_id, method, provider_message_id)"
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_message_cache_folder ON message_cache(mailbox_id, method, folder_id)"
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_message_cache_conversation ON message_cache(mailbox_id, method, conversation_id)"
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_message_meta_lookup ON message_meta(mailbox_id, method, provider_message_id)"
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_attachment_cache_lookup ON attachment_cache(mailbox_id, method, provider_message_id)"
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_audit_logs_lookup ON audit_logs(mailbox_id, action, created_at DESC)"
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_sync_jobs_lookup ON sync_jobs(mailbox_id, method, started_at DESC)"
            )
            try:
                connection.execute(
                    """
                    CREATE VIRTUAL TABLE IF NOT EXISTS message_search USING fts5(
                        document_id UNINDEXED,
                        subject,
                        sender,
                        preview,
                        body_text,
                        notes,
                        tags_text
                    )
                    """
                )
                self._fts_enabled = True
            except sqlite3.OperationalError:
                self._fts_enabled = False
                connection.execute(
                    """
                    CREATE TABLE IF NOT EXISTS message_search (
                        document_id TEXT PRIMARY KEY,
                        subject TEXT NOT NULL DEFAULT '',
                        sender TEXT NOT NULL DEFAULT '',
                        preview TEXT NOT NULL DEFAULT '',
                        body_text TEXT NOT NULL DEFAULT '',
                        notes TEXT NOT NULL DEFAULT '',
                        tags_text TEXT NOT NULL DEFAULT ''
                    )
                    """
                )

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
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

        email = self._normalize_text(payload.get("email"), fallback=current.email if current else None)
        client_id = self._normalize_text(payload.get("client_id"), fallback=current.client_id if current else None)
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

    def _normalize_method_value(self, method: str | None) -> str:
        normalized = self._normalize_optional_string(method) or DEFAULT_METHOD
        if normalized not in VALID_METHODS:
            raise MailboxStoreError("preferred_method 仅支持 graph_api、imap_new、imap_old")
        return normalized

    def _normalize_optional_string(self, value: Any) -> str:
        if value is None:
            return ""
        if not isinstance(value, str):
            raise MailboxStoreError("字段必须是字符串")
        return value.strip()

    def _require_non_empty_text(self, value: Any, message: str) -> str:
        normalized = self._normalize_optional_string(value)
        if not normalized:
            raise MailboxStoreError(message)
        return normalized

    def _normalize_tags(self, tags: list[str] | None) -> list[str]:
        if tags is None:
            return []
        if not isinstance(tags, list):
            raise MailboxStoreError("tags 必须是字符串数组")
        normalized: list[str] = []
        seen: set[str] = set()
        for item in tags:
            if not isinstance(item, str):
                raise MailboxStoreError("tags 中的每一项都必须是字符串")
            cleaned = item.strip()
            if not cleaned:
                continue
            marker = cleaned.casefold()
            if marker in seen:
                continue
            seen.add(marker)
            normalized.append(cleaned)
        return normalized

    @staticmethod
    def _json_dumps(value: Any) -> str:
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))

    def _build_document_id(self, mailbox_id: int, method: str, message_id: str) -> str:
        return f"{mailbox_id}|{method}|{message_id}"

    def _build_fts_query(self, query: str) -> str:
        tokens = [token.strip() for token in query.replace('"', " ").split() if token.strip()]
        if not tokens:
            return ""
        return " AND ".join(f"{token}*" for token in tokens)

    def _normalize_folder_record(
        self,
        mailbox_id: int,
        method: str,
        item: Any,
        now: str,
    ) -> FolderCacheEntry:
        source = asdict(item) if is_dataclass(item) else dict(item or {})
        folder_id = self._require_non_empty_text(source.get("id") or source.get("folder_id") or source.get("name"), "缺少文件夹标识")
        display_name = self._normalize_optional_string(source.get("display_name")) or folder_id
        return FolderCacheEntry(
            mailbox_id=mailbox_id,
            method=method,
            folder_id=folder_id,
            name=self._normalize_optional_string(source.get("name")) or display_name,
            display_name=display_name,
            kind=self._normalize_optional_string(source.get("kind")) or "custom",
            total=int(source.get("total", 0) or 0),
            unread=int(source.get("unread", 0) or 0),
            is_default=bool(source.get("is_default", False)),
            cached_at=now,
            updated_at=now,
        )

    def _normalize_message_record(
        self,
        mailbox_id: int,
        method: str,
        item: Any,
        now: str,
    ) -> dict[str, Any]:
        source = asdict(item) if is_dataclass(item) else dict(item or {})
        provider_message_id = self._require_non_empty_text(
            source.get("message_id") or source.get("id") or source.get("provider_message_id"),
            "缺少邮件标识",
        )
        attachments = source.get("attachments")
        attachments_list = list(attachments) if isinstance(attachments, list) else []
        raw_json = self._json_dumps(source if isinstance(source, dict) else {})
        preview = self._normalize_optional_string(source.get("preview"))
        body_text = self._normalize_optional_string(source.get("body_text")) or self._normalize_optional_string(source.get("body"))
        return {
            "mailbox_id": mailbox_id,
            "method": method,
            "provider_message_id": provider_message_id,
            "internet_message_id": self._normalize_optional_string(source.get("internet_message_id")),
            "conversation_id": self._normalize_optional_string(source.get("conversation_id")),
            "folder_id": self._normalize_optional_string(source.get("folder")) or self._normalize_optional_string(source.get("folder_id")),
            "subject": self._normalize_optional_string(source.get("subject")) or "无主题",
            "sender": self._normalize_optional_string(source.get("sender")) or "未知发件人",
            "sender_name": self._normalize_optional_string(source.get("sender_name")),
            "received_at": self._normalize_optional_string(source.get("received_at")),
            "is_read": int(bool(source.get("is_read", False))),
            "is_flagged": int(bool(source.get("is_flagged", False))),
            "importance": self._normalize_optional_string(source.get("importance")) or "normal",
            "has_attachments": int(bool(source.get("has_attachments", False) or attachments_list)),
            "preview": preview or body_text[:240],
            "body_text": body_text,
            "body_html": self._normalize_optional_string(source.get("body_html")),
            "to_recipients_json": self._json_dumps(list(source.get("to_recipients", source.get("to", [])) or [])),
            "cc_recipients_json": self._json_dumps(list(source.get("cc_recipients", source.get("cc", [])) or [])),
            "bcc_recipients_json": self._json_dumps(list(source.get("bcc_recipients", source.get("bcc", [])) or [])),
            "headers_json": self._json_dumps(dict(source.get("headers", {}) or {})),
            "in_reply_to": self._normalize_optional_string(source.get("in_reply_to")),
            "references_json": self._json_dumps(list(source.get("references", []) or [])),
            "raw_json": raw_json,
            "attachments": attachments_list,
            "attachments_provided": isinstance(attachments, list),
            "cached_at": now,
            "updated_at": now,
        }

    def _normalize_attachment_record(
        self,
        mailbox_id: int,
        method: str,
        message_id: str,
        attachment: Any,
        now: str,
    ) -> CachedAttachment:
        source = asdict(attachment) if is_dataclass(attachment) else dict(attachment or {})
        attachment_id = self._require_non_empty_text(
            source.get("attachment_id") or source.get("id"),
            "缺少附件标识",
        )
        return CachedAttachment(
            mailbox_id=mailbox_id,
            method=method,
            message_id=message_id,
            attachment_id=attachment_id,
            name=self._normalize_optional_string(source.get("name")) or attachment_id,
            content_type=self._normalize_optional_string(source.get("content_type")) or "application/octet-stream",
            size=int(source.get("size", 0) or 0),
            is_inline=bool(source.get("is_inline", False)),
            content_base64=self._normalize_optional_string(source.get("content_base64")),
            cached_at=now,
            updated_at=now,
        )

    def _replace_attachment_cache(
        self,
        connection: sqlite3.Connection,
        *,
        mailbox_id: int,
        method: str,
        message_id: str,
        attachments: Iterable[Any],
        now: str,
    ) -> None:
        connection.execute(
            "DELETE FROM attachment_cache WHERE mailbox_id = ? AND method = ? AND provider_message_id = ?",
            (mailbox_id, method, message_id),
        )
        records = [
            self._normalize_attachment_record(mailbox_id, method, message_id, item, now)
            for item in attachments
        ]
        if not records:
            return
        connection.executemany(
            """
            INSERT INTO attachment_cache (
                mailbox_id,
                method,
                provider_message_id,
                attachment_id,
                name,
                content_type,
                size,
                is_inline,
                content_base64,
                cached_at,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    record.mailbox_id,
                    record.method,
                    record.message_id,
                    record.attachment_id,
                    record.name,
                    record.content_type,
                    record.size,
                    int(record.is_inline),
                    record.content_base64,
                    record.cached_at,
                    record.updated_at,
                )
                for record in records
            ],
        )

    def _refresh_search_document(
        self,
        connection: sqlite3.Connection,
        mailbox_id: int,
        method: str,
        message_id: str,
    ) -> None:
        row = connection.execute(
            """
            SELECT
                c.subject,
                c.sender,
                c.preview,
                c.body_text,
                COALESCE(mm.notes, '') AS notes,
                COALESCE(mm.tags_text, '') AS tags_text
            FROM message_cache AS c
            LEFT JOIN message_meta AS mm
                ON mm.mailbox_id = c.mailbox_id
                AND mm.method = c.method
                AND mm.provider_message_id = c.provider_message_id
            WHERE c.mailbox_id = ? AND c.method = ? AND c.provider_message_id = ?
            LIMIT 1
            """,
            (mailbox_id, method, message_id),
        ).fetchone()
        document_id = self._build_document_id(mailbox_id, method, message_id)
        if not row:
            if self._fts_enabled:
                connection.execute("DELETE FROM message_search WHERE document_id = ?", (document_id,))
            return

        connection.execute("DELETE FROM message_search WHERE document_id = ?", (document_id,))
        connection.execute(
            """
            INSERT INTO message_search (
                document_id,
                subject,
                sender,
                preview,
                body_text,
                notes,
                tags_text
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                document_id,
                str(row["subject"] or ""),
                str(row["sender"] or ""),
                str(row["preview"] or ""),
                str(row["body_text"] or ""),
                str(row["notes"] or ""),
                str(row["tags_text"] or ""),
            ),
        )

    def _list_attachments_from_connection(
        self,
        connection: sqlite3.Connection,
        *,
        mailbox_id: int,
        method: str,
        message_id: str,
    ) -> list[CachedAttachment]:
        rows = connection.execute(
            """
            SELECT
                mailbox_id,
                method,
                provider_message_id,
                attachment_id,
                name,
                content_type,
                size,
                is_inline,
                content_base64,
                cached_at,
                updated_at
            FROM attachment_cache
            WHERE mailbox_id = ? AND method = ? AND provider_message_id = ?
            ORDER BY attachment_id ASC
            """,
            (mailbox_id, method, message_id),
        ).fetchall()
        return [self._row_to_attachment(row) for row in rows]

    def _list_attachments_map_from_connection(
        self,
        connection: sqlite3.Connection,
        *,
        mailbox_id: int,
        method: str | None,
        message_ids: list[str],
    ) -> dict[str, list[CachedAttachment]]:
        unique_ids = [self._normalize_optional_string(item) for item in dict.fromkeys(message_ids)]
        unique_ids = [item for item in unique_ids if item]
        if not unique_ids:
            return {}
        placeholders = ", ".join("?" for _ in unique_ids)
        conditions = ["mailbox_id = ?"]
        parameters: list[Any] = [mailbox_id]
        if method:
            conditions.append("method = ?")
            parameters.append(method)
        conditions.append(f"provider_message_id IN ({placeholders})")
        parameters.extend(unique_ids)
        rows = connection.execute(
            f"""
            SELECT
                mailbox_id,
                method,
                provider_message_id,
                attachment_id,
                name,
                content_type,
                size,
                is_inline,
                content_base64,
                cached_at,
                updated_at
            FROM attachment_cache
            WHERE {' AND '.join(conditions)}
            ORDER BY provider_message_id ASC, attachment_id ASC
            """,
            parameters,
        ).fetchall()
        mapped: dict[str, list[CachedAttachment]] = {}
        for row in rows:
            mapped.setdefault(str(row["provider_message_id"]), []).append(self._row_to_attachment(row))
        return mapped

    def _row_to_profile(self, row: sqlite3.Row) -> MailboxProfile:
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

    def _row_to_summary(self, row: sqlite3.Row) -> MailboxSummary:
        return MailboxSummary(
            id=int(row["id"]),
            label=str(row["label"]),
            email=str(row["email"]),
            preferred_method=str(row["preferred_method"]),
            notes=str(row["notes"] or ""),
            created_at=str(row["created_at"]),
            updated_at=str(row["updated_at"]),
        )

    def _row_to_folder_entry(self, row: sqlite3.Row) -> FolderCacheEntry:
        return FolderCacheEntry(
            mailbox_id=int(row["mailbox_id"]),
            method=str(row["method"]),
            folder_id=str(row["folder_id"]),
            name=str(row["name"]),
            display_name=str(row["display_name"]),
            kind=str(row["kind"] or "custom"),
            total=int(row["total"] or 0),
            unread=int(row["unread"] or 0),
            is_default=bool(row["is_default"]),
            cached_at=str(row["cached_at"]),
            updated_at=str(row["updated_at"]),
        )

    def _row_to_message_meta(self, row: sqlite3.Row) -> MessageMeta:
        tags = self._load_json_list(row["tags_json"])
        return MessageMeta(
            mailbox_id=int(row["mailbox_id"]),
            method=str(row["method"]),
            message_id=str(row["provider_message_id"]),
            tags=tags,
            follow_up=str(row["follow_up"] or ""),
            notes=str(row["notes"] or ""),
            snoozed_until=str(row["snoozed_until"] or ""),
            status=str(row["status"] or "active"),
            updated_at=str(row["updated_at"]),
        )

    def _row_to_attachment(self, row: sqlite3.Row) -> CachedAttachment:
        return CachedAttachment(
            mailbox_id=int(row["mailbox_id"]),
            method=str(row["method"]),
            message_id=str(row["provider_message_id"]),
            attachment_id=str(row["attachment_id"]),
            name=str(row["name"]),
            content_type=str(row["content_type"] or "application/octet-stream"),
            size=int(row["size"] or 0),
            is_inline=bool(row["is_inline"]),
            content_base64=str(row["content_base64"] or ""),
            cached_at=str(row["cached_at"]),
            updated_at=str(row["updated_at"]),
        )

    def _row_to_cached_message(
        self,
        row: sqlite3.Row,
        *,
        attachments: list[CachedAttachment],
    ) -> CachedMessage:
        meta = None
        if row["meta_updated_at"]:
            meta = MessageMeta(
                mailbox_id=int(row["mailbox_id"]),
                method=str(row["method"]),
                message_id=str(row["provider_message_id"]),
                tags=self._load_json_list(row["tags_json"]),
                follow_up=str(row["follow_up"] or ""),
                notes=str(row["notes"] or ""),
                snoozed_until=str(row["snoozed_until"] or ""),
                status=str(row["meta_status"] or "active"),
                updated_at=str(row["meta_updated_at"]),
            )
        return CachedMessage(
            mailbox_id=int(row["mailbox_id"]),
            mailbox_label=str(row["mailbox_label"]),
            mailbox_email=str(row["mailbox_email"]),
            method=str(row["method"]),
            message_id=str(row["provider_message_id"]),
            subject=str(row["subject"] or "无主题"),
            sender=str(row["sender"] or "未知发件人"),
            sender_name=str(row["sender_name"] or ""),
            received_at=str(row["received_at"] or ""),
            is_read=bool(row["is_read"]),
            is_flagged=bool(row["is_flagged"]),
            importance=str(row["importance"] or "normal"),
            has_attachments=bool(row["has_attachments"]),
            preview=str(row["preview"] or ""),
            body_text=str(row["body_text"] or ""),
            body_html=str(row["body_html"]) if row["body_html"] else None,
            folder=str(row["folder_id"] or ""),
            internet_message_id=str(row["internet_message_id"] or ""),
            conversation_id=str(row["conversation_id"] or ""),
            to_recipients=self._load_json_list(row["to_recipients_json"]),
            cc_recipients=self._load_json_list(row["cc_recipients_json"]),
            bcc_recipients=self._load_json_list(row["bcc_recipients_json"]),
            headers=self._load_json_dict(row["headers_json"]),
            in_reply_to=str(row["in_reply_to"] or ""),
            references=self._load_json_list(row["references_json"]),
            attachments=attachments,
            meta=meta,
            cached_at=str(row["cached_at"]),
            updated_at=str(row["updated_at"]),
        )

    def _row_to_search_result(self, row: sqlite3.Row) -> MessageSearchResult:
        meta = None
        if row["meta_updated_at"]:
            meta = MessageMeta(
                mailbox_id=int(row["mailbox_id"]),
                method=str(row["method"]),
                message_id=str(row["provider_message_id"]),
                tags=self._load_json_list(row["tags_json"]),
                follow_up=str(row["follow_up"] or ""),
                notes=str(row["notes"] or ""),
                snoozed_until=str(row["snoozed_until"] or ""),
                status=str(row["meta_status"] or "active"),
                updated_at=str(row["meta_updated_at"]),
            )
        return MessageSearchResult(
            mailbox_id=int(row["mailbox_id"]),
            mailbox_label=str(row["mailbox_label"]),
            mailbox_email=str(row["mailbox_email"]),
            method=str(row["method"]),
            message_id=str(row["provider_message_id"]),
            subject=str(row["subject"] or "无主题"),
            sender=str(row["sender"] or "未知发件人"),
            sender_name=str(row["sender_name"] or ""),
            received_at=str(row["received_at"] or ""),
            is_read=bool(row["is_read"]),
            is_flagged=bool(row["is_flagged"]),
            importance=str(row["importance"] or "normal"),
            has_attachments=bool(row["has_attachments"]),
            preview=str(row["preview"] or ""),
            folder=str(row["folder_id"] or ""),
            internet_message_id=str(row["internet_message_id"] or ""),
            conversation_id=str(row["conversation_id"] or ""),
            meta=meta,
        )

    def _row_to_rule(self, row: sqlite3.Row) -> SavedRule:
        return SavedRule(
            id=int(row["id"]),
            mailbox_id=int(row["mailbox_id"]),
            name=str(row["name"]),
            enabled=bool(row["enabled"]),
            priority=int(row["priority"] or 0),
            conditions=self._load_json_dict(row["conditions_json"]),
            actions=self._load_json_dict(row["actions_json"]),
            created_at=str(row["created_at"]),
            updated_at=str(row["updated_at"]),
        )

    def _row_to_audit_log(self, row: sqlite3.Row) -> AuditLogEntry:
        return AuditLogEntry(
            id=int(row["id"]),
            mailbox_id=int(row["mailbox_id"]) if row["mailbox_id"] is not None else None,
            mailbox_label=str(row["mailbox_label"] or ""),
            mailbox_email=str(row["mailbox_email"] or ""),
            actor=str(row["actor"] or "system"),
            action=str(row["action"] or ""),
            target_type=str(row["target_type"] or ""),
            target_id=str(row["target_id"] or ""),
            status=str(row["status"] or "success"),
            details=self._load_json_dict(row["details_json"]),
            created_at=str(row["created_at"]),
        )

    def _row_to_sync_job(self, row: sqlite3.Row) -> SyncJob:
        return SyncJob(
            id=int(row["id"]),
            mailbox_id=int(row["mailbox_id"]),
            method=str(row["method"]),
            requested_by=str(row["requested_by"] or "system"),
            scope=self._load_json_dict(row["scope_json"]),
            status=str(row["status"] or "running"),
            processed_messages=int(row["processed_messages"] or 0),
            cached_messages=int(row["cached_messages"] or 0),
            folders_synced=int(row["folders_synced"] or 0),
            error=str(row["error"] or ""),
            started_at=str(row["started_at"]),
            finished_at=str(row["finished_at"] or ""),
        )

    def _row_to_sync_state(self, row: sqlite3.Row) -> SyncState:
        return SyncState(
            mailbox_id=int(row["mailbox_id"]),
            method=str(row["method"]),
            folder_id=str(row["folder_id"]),
            last_synced_at=str(row["last_synced_at"] or ""),
            last_message_at=str(row["last_message_at"] or ""),
            cached_messages=int(row["cached_messages"] or 0),
            status=str(row["status"] or "idle"),
            error=str(row["error"] or ""),
            updated_at=str(row["updated_at"]),
        )

    @staticmethod
    def _load_json_list(raw: Any) -> list[Any]:
        if raw in (None, ""):
            return []
        if isinstance(raw, list):
            return raw
        try:
            value = json.loads(str(raw))
        except json.JSONDecodeError:
            return []
        return value if isinstance(value, list) else []

    @staticmethod
    def _load_json_dict(raw: Any) -> dict[str, Any]:
        if raw in (None, ""):
            return {}
        if isinstance(raw, dict):
            return raw
        try:
            value = json.loads(str(raw))
        except json.JSONDecodeError:
            return {}
        return value if isinstance(value, dict) else {}

    @staticmethod
    def _normalize_text(value: Any, fallback: str | None = None) -> str | None:
        if value is None:
            return fallback
        if not isinstance(value, str):
            raise MailboxStoreError("邮箱档案字段必须是字符串")
        cleaned = value.strip()
        return cleaned or None

    @staticmethod
    def _utc_now() -> str:
        return datetime.now(UTC).isoformat(timespec="seconds").replace("+00:00", "Z")

    def _initialize_schema(self) -> None:
        with self._connect() as connection:
            connection.executescript(
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
                );

                CREATE TABLE IF NOT EXISTS folder_cache (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    folder_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    display_name TEXT NOT NULL,
                    kind TEXT NOT NULL DEFAULT 'custom',
                    total INTEGER NOT NULL DEFAULT 0,
                    unread INTEGER NOT NULL DEFAULT 0,
                    is_default INTEGER NOT NULL DEFAULT 0,
                    cached_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(mailbox_id, method, folder_id),
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS message_cache (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    provider_message_id TEXT NOT NULL,
                    internet_message_id TEXT NOT NULL DEFAULT '',
                    conversation_id TEXT NOT NULL DEFAULT '',
                    folder_id TEXT NOT NULL DEFAULT '',
                    subject TEXT NOT NULL DEFAULT '',
                    sender TEXT NOT NULL DEFAULT '',
                    sender_name TEXT NOT NULL DEFAULT '',
                    received_at TEXT NOT NULL DEFAULT '',
                    is_read INTEGER NOT NULL DEFAULT 0,
                    is_flagged INTEGER NOT NULL DEFAULT 0,
                    importance TEXT NOT NULL DEFAULT 'normal',
                    has_attachments INTEGER NOT NULL DEFAULT 0,
                    preview TEXT NOT NULL DEFAULT '',
                    body_text TEXT NOT NULL DEFAULT '',
                    body_html TEXT NOT NULL DEFAULT '',
                    to_recipients_json TEXT NOT NULL DEFAULT '[]',
                    cc_recipients_json TEXT NOT NULL DEFAULT '[]',
                    bcc_recipients_json TEXT NOT NULL DEFAULT '[]',
                    headers_json TEXT NOT NULL DEFAULT '{}',
                    in_reply_to TEXT NOT NULL DEFAULT '',
                    references_json TEXT NOT NULL DEFAULT '[]',
                    raw_json TEXT NOT NULL DEFAULT '{}',
                    cached_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(mailbox_id, method, provider_message_id),
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS attachment_cache (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    provider_message_id TEXT NOT NULL,
                    attachment_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    content_type TEXT NOT NULL DEFAULT '',
                    size INTEGER NOT NULL DEFAULT 0,
                    is_inline INTEGER NOT NULL DEFAULT 0,
                    content_base64 TEXT NOT NULL DEFAULT '',
                    cached_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    UNIQUE(mailbox_id, method, provider_message_id, attachment_id),
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS message_meta (
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    provider_message_id TEXT NOT NULL,
                    tags_json TEXT NOT NULL DEFAULT '[]',
                    tags_text TEXT NOT NULL DEFAULT '',
                    follow_up TEXT NOT NULL DEFAULT '',
                    notes TEXT NOT NULL DEFAULT '',
                    snoozed_until TEXT NOT NULL DEFAULT '',
                    status TEXT NOT NULL DEFAULT 'active',
                    updated_at TEXT NOT NULL,
                    PRIMARY KEY(mailbox_id, method, provider_message_id),
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS saved_rules (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    enabled INTEGER NOT NULL DEFAULT 1,
                    priority INTEGER NOT NULL DEFAULT 100,
                    conditions_json TEXT NOT NULL DEFAULT '{}',
                    actions_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS audit_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER,
                    actor TEXT NOT NULL,
                    action TEXT NOT NULL,
                    target_type TEXT NOT NULL,
                    target_id TEXT NOT NULL DEFAULT '',
                    status TEXT NOT NULL DEFAULT 'success',
                    details_json TEXT NOT NULL DEFAULT '{}',
                    created_at TEXT NOT NULL,
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE SET NULL
                );

                CREATE TABLE IF NOT EXISTS sync_jobs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    requested_by TEXT NOT NULL,
                    scope_json TEXT NOT NULL DEFAULT '{}',
                    status TEXT NOT NULL DEFAULT 'queued',
                    processed_messages INTEGER NOT NULL DEFAULT 0,
                    cached_messages INTEGER NOT NULL DEFAULT 0,
                    folders_synced INTEGER NOT NULL DEFAULT 0,
                    error TEXT NOT NULL DEFAULT '',
                    started_at TEXT NOT NULL,
                    finished_at TEXT NOT NULL DEFAULT '',
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS sync_state (
                    mailbox_id INTEGER NOT NULL,
                    method TEXT NOT NULL,
                    folder_id TEXT NOT NULL,
                    last_synced_at TEXT NOT NULL,
                    last_message_at TEXT NOT NULL DEFAULT '',
                    cached_messages INTEGER NOT NULL DEFAULT 0,
                    status TEXT NOT NULL DEFAULT 'idle',
                    error TEXT NOT NULL DEFAULT '',
                    updated_at TEXT NOT NULL,
                    PRIMARY KEY(mailbox_id, method, folder_id),
                    FOREIGN KEY(mailbox_id) REFERENCES mailboxes(id) ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS idx_folder_cache_mailbox_method
                    ON folder_cache(mailbox_id, method);
                CREATE INDEX IF NOT EXISTS idx_message_cache_mailbox_method_received
                    ON message_cache(mailbox_id, method, received_at DESC);
                CREATE INDEX IF NOT EXISTS idx_message_cache_conversation
                    ON message_cache(mailbox_id, conversation_id);
                CREATE INDEX IF NOT EXISTS idx_message_cache_internet_message
                    ON message_cache(mailbox_id, internet_message_id);
                CREATE INDEX IF NOT EXISTS idx_message_cache_folder
                    ON message_cache(mailbox_id, method, folder_id);
                CREATE INDEX IF NOT EXISTS idx_attachment_cache_message
                    ON attachment_cache(mailbox_id, method, provider_message_id);
                CREATE INDEX IF NOT EXISTS idx_saved_rules_mailbox
                    ON saved_rules(mailbox_id, enabled, priority);
                CREATE INDEX IF NOT EXISTS idx_audit_logs_mailbox_created
                    ON audit_logs(mailbox_id, created_at DESC);
                CREATE INDEX IF NOT EXISTS idx_sync_jobs_mailbox_started
                    ON sync_jobs(mailbox_id, started_at DESC);
                """
            )
            try:
                connection.execute(
                    """
                    CREATE VIRTUAL TABLE IF NOT EXISTS message_search USING fts5(
                        document_id UNINDEXED,
                        mailbox_id UNINDEXED,
                        method UNINDEXED,
                        provider_message_id UNINDEXED,
                        folder_id UNINDEXED,
                        subject,
                        sender,
                        preview,
                        body_text,
                        tags,
                        notes,
                        tokenize = 'unicode61'
                    )
                    """
                )
            except sqlite3.OperationalError:
                self._fts_enabled = False

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
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
            "label": label or "",
            "email": email or "",
            "client_id": client_id or "",
            "refresh_token": refresh_token or "",
            "proxy": proxy,
            "preferred_method": preferred_method,
            "notes": notes,
        }

    def _normalize_folder_record(
        self,
        mailbox_id: int,
        method: str,
        raw: Any,
        now: str,
    ) -> FolderCacheEntry:
        item = self._coerce_mapping(raw)
        folder_id = self._normalize_optional_string(item.get("id")) or self._normalize_optional_string(item.get("name")) or "custom"
        display_name = self._normalize_optional_string(item.get("display_name")) or folder_id
        name = self._normalize_optional_string(item.get("name")) or display_name
        return FolderCacheEntry(
            mailbox_id=mailbox_id,
            method=method,
            folder_id=folder_id,
            name=name,
            display_name=display_name,
            kind=self._normalize_optional_string(item.get("kind")) or "custom",
            total=self._coerce_int(item.get("total"), default=0),
            unread=self._coerce_int(item.get("unread"), default=0),
            is_default=bool(item.get("is_default", False)),
            cached_at=now,
            updated_at=now,
        )

    def _normalize_message_record(
        self,
        mailbox_id: int,
        method: str,
        raw: Any,
        now: str,
    ) -> dict[str, Any]:
        item = self._coerce_mapping(raw)
        provider_message_id = self._normalize_optional_string(item.get("message_id") or item.get("id"))
        if not provider_message_id:
            raise MailboxStoreError("消息缓存缺少 provider message id")

        headers = self._coerce_headers(item.get("headers"))
        attachments_provided = "attachments" in item
        attachments = item.get("attachments") if attachments_provided else []
        to_recipients = self._coerce_string_list(item.get("to_recipients") or item.get("to"))
        cc_recipients = self._coerce_string_list(item.get("cc_recipients") or item.get("cc"))
        bcc_recipients = self._coerce_string_list(item.get("bcc_recipients") or item.get("bcc"))
        references = self._parse_reference_header(headers.get("References", ""))
        body_text = self._normalize_optional_string(item.get("body_text") or item.get("body")) or ""
        preview = self._normalize_optional_string(item.get("preview")) or (body_text[:180] if body_text else "")

        return {
            "mailbox_id": mailbox_id,
            "method": method,
            "provider_message_id": provider_message_id,
            "internet_message_id": self._normalize_optional_string(item.get("internet_message_id")) or "",
            "conversation_id": self._normalize_optional_string(item.get("conversation_id")) or "",
            "folder_id": self._normalize_optional_string(item.get("folder")) or "",
            "subject": self._normalize_optional_string(item.get("subject")) or "",
            "sender": self._normalize_optional_string(item.get("sender")) or "",
            "sender_name": self._normalize_optional_string(item.get("sender_name")) or "",
            "received_at": self._normalize_optional_string(item.get("received_at")) or "",
            "is_read": int(bool(item.get("is_read", False))),
            "is_flagged": int(bool(item.get("is_flagged", False))),
            "importance": self._normalize_optional_string(item.get("importance")) or "normal",
            "has_attachments": int(bool(item.get("has_attachments", False) or attachments)),
            "preview": preview,
            "body_text": body_text,
            "body_html": self._normalize_optional_string(item.get("body_html")) or "",
            "to_recipients_json": self._json_dumps(to_recipients),
            "cc_recipients_json": self._json_dumps(cc_recipients),
            "bcc_recipients_json": self._json_dumps(bcc_recipients),
            "headers_json": self._json_dumps(headers),
            "in_reply_to": self._normalize_optional_string(headers.get("In-Reply-To")) or "",
            "references_json": self._json_dumps(references),
            "raw_json": self._json_dumps(self._jsonable(item)),
            "cached_at": now,
            "updated_at": now,
            "attachments": list(attachments) if isinstance(attachments, list) else [],
            "attachments_provided": attachments_provided,
        }

    def _replace_attachment_cache(
        self,
        connection: sqlite3.Connection,
        *,
        mailbox_id: int,
        method: str,
        message_id: str,
        attachments: Iterable[Any],
        now: str,
    ) -> None:
        connection.execute(
            "DELETE FROM attachment_cache WHERE mailbox_id = ? AND method = ? AND provider_message_id = ?",
            (mailbox_id, method, message_id),
        )
        records = [
            self._normalize_attachment_record(mailbox_id, method, message_id, item, now)
            for item in attachments
        ]
        if not records:
            return
        connection.executemany(
            """
            INSERT INTO attachment_cache (
                mailbox_id,
                method,
                provider_message_id,
                attachment_id,
                name,
                content_type,
                size,
                is_inline,
                content_base64,
                cached_at,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    record.mailbox_id,
                    record.method,
                    record.message_id,
                    record.attachment_id,
                    record.name,
                    record.content_type,
                    record.size,
                    int(record.is_inline),
                    record.content_base64,
                    record.cached_at,
                    record.updated_at,
                )
                for record in records
            ],
        )

    def _normalize_attachment_record(
        self,
        mailbox_id: int,
        method: str,
        message_id: str,
        raw: Any,
        now: str,
    ) -> CachedAttachment:
        item = self._coerce_mapping(raw)
        attachment_id = self._normalize_optional_string(item.get("attachment_id") or item.get("id"))
        if not attachment_id:
            raise MailboxStoreError("附件缓存缺少 attachment id")
        return CachedAttachment(
            mailbox_id=mailbox_id,
            method=method,
            message_id=message_id,
            attachment_id=attachment_id,
            name=self._normalize_optional_string(item.get("name")) or attachment_id,
            content_type=self._normalize_optional_string(item.get("content_type")) or "",
            size=self._coerce_int(item.get("size"), default=0),
            is_inline=bool(item.get("is_inline", False)),
            content_base64=self._normalize_optional_string(item.get("content_base64")) or "",
            cached_at=now,
            updated_at=now,
        )

    def _refresh_search_document(
        self,
        connection: sqlite3.Connection,
        mailbox_id: int,
        method: str,
        message_id: str,
    ) -> None:
        if not self._fts_enabled:
            return

        document_id = self._build_document_id(mailbox_id, method, message_id)
        connection.execute("DELETE FROM message_search WHERE document_id = ?", (document_id,))
        row = connection.execute(
            """
            SELECT
                c.mailbox_id,
                c.method,
                c.provider_message_id,
                c.folder_id,
                c.subject,
                c.sender,
                c.preview,
                c.body_text,
                COALESCE(mm.tags_text, '') AS tags_text,
                COALESCE(mm.notes, '') AS notes
            FROM message_cache AS c
            LEFT JOIN message_meta AS mm
                ON mm.mailbox_id = c.mailbox_id
                AND mm.method = c.method
                AND mm.provider_message_id = c.provider_message_id
            WHERE c.mailbox_id = ? AND c.method = ? AND c.provider_message_id = ?
            LIMIT 1
            """,
            (mailbox_id, method, message_id),
        ).fetchone()
        if not row:
            return
        connection.execute(
            """
            INSERT INTO message_search (
                document_id,
                mailbox_id,
                method,
                provider_message_id,
                folder_id,
                subject,
                sender,
                preview,
                body_text,
                tags,
                notes
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                document_id,
                str(row["mailbox_id"]),
                str(row["method"]),
                str(row["provider_message_id"]),
                str(row["folder_id"] or ""),
                str(row["subject"] or ""),
                str(row["sender"] or ""),
                str(row["preview"] or ""),
                str(row["body_text"] or ""),
                str(row["tags_text"] or ""),
                str(row["notes"] or ""),
            ),
        )

    def _list_attachments_from_connection(
        self,
        connection: sqlite3.Connection,
        *,
        mailbox_id: int,
        method: str,
        message_id: str,
    ) -> list[CachedAttachment]:
        rows = connection.execute(
            """
            SELECT
                mailbox_id,
                method,
                provider_message_id,
                attachment_id,
                name,
                content_type,
                size,
                is_inline,
                content_base64,
                cached_at,
                updated_at
            FROM attachment_cache
            WHERE mailbox_id = ? AND method = ? AND provider_message_id = ?
            ORDER BY id ASC
            """,
            (mailbox_id, method, message_id),
        ).fetchall()
        return [self._row_to_attachment(row) for row in rows]

    def _list_attachments_map_from_connection(
        self,
        connection: sqlite3.Connection,
        *,
        mailbox_id: int,
        method: str | None,
        message_ids: list[str],
    ) -> dict[str, list[CachedAttachment]]:
        unique_ids = [item for item in dict.fromkeys(message_ids) if item]
        if not unique_ids:
            return {}
        placeholders = ", ".join("?" for _ in unique_ids)
        parameters: list[Any] = [mailbox_id]
        method_clause = ""
        if method:
            method_clause = "AND method = ?"
            parameters.append(method)
        rows = connection.execute(
            f"""
            SELECT
                mailbox_id,
                method,
                provider_message_id,
                attachment_id,
                name,
                content_type,
                size,
                is_inline,
                content_base64,
                cached_at,
                updated_at
            FROM attachment_cache
            WHERE mailbox_id = ?
              {method_clause}
              AND provider_message_id IN ({placeholders})
            ORDER BY id ASC
            """,
            [*parameters, *unique_ids],
        ).fetchall()
        grouped: dict[str, list[CachedAttachment]] = {}
        for row in rows:
            grouped.setdefault(str(row["provider_message_id"]), []).append(self._row_to_attachment(row))
        return grouped

    @staticmethod
    def _normalize_text(value: Any, fallback: str | None = None) -> str | None:
        if value is None:
            return fallback
        if not isinstance(value, str):
            raise MailboxStoreError("邮箱档案字段必须是字符串")
        cleaned = value.strip()
        return cleaned or None

    @staticmethod
    def _normalize_optional_string(value: Any) -> str | None:
        if value is None:
            return None
        if not isinstance(value, str):
            raise MailboxStoreError("请求字段必须是字符串")
        cleaned = value.strip()
        return cleaned or None

    @staticmethod
    def _require_non_empty_text(value: Any, message: str) -> str:
        if not isinstance(value, str) or not value.strip():
            raise MailboxStoreError(message)
        return value.strip()

    @staticmethod
    def _normalize_method_value(method: str) -> str:
        normalized = str(method or "").strip()
        if normalized not in VALID_METHODS:
            raise MailboxStoreError("不支持的邮件接入方式")
        return normalized

    @staticmethod
    def _normalize_tags(value: list[str] | None) -> list[str]:
        if value is None:
            return []
        if not isinstance(value, list):
            raise MailboxStoreError("tags 必须是数组")
        tags: list[str] = []
        seen: set[str] = set()
        for item in value:
            if not isinstance(item, str) or not item.strip():
                raise MailboxStoreError("tags 中的每一项都必须是非空字符串")
            normalized = item.strip()
            key = normalized.casefold()
            if key in seen:
                continue
            seen.add(key)
            tags.append(normalized)
        return tags

    @staticmethod
    def _coerce_int(value: Any, *, default: int) -> int:
        if value in (None, ""):
            return default
        try:
            return int(value)
        except (TypeError, ValueError) as exc:
            raise MailboxStoreError("请求字段必须是整数") from exc

    def _coerce_mapping(self, value: Any) -> dict[str, Any]:
        if is_dataclass(value):
            return asdict(value)
        if isinstance(value, dict):
            return dict(value)
        raise MailboxStoreError("缓存对象必须是 dataclass 或 dict")

    @staticmethod
    def _coerce_string_list(value: Any) -> list[str]:
        if value is None:
            return []
        if not isinstance(value, list):
            raise MailboxStoreError("收件人字段必须是数组")
        items: list[str] = []
        for item in value:
            if not isinstance(item, str):
                raise MailboxStoreError("收件人列表中的每一项都必须是字符串")
            cleaned = item.strip()
            if cleaned:
                items.append(cleaned)
        return items

    @staticmethod
    def _coerce_headers(value: Any) -> dict[str, str]:
        if value is None:
            return {}
        if not isinstance(value, dict):
            raise MailboxStoreError("headers 必须是对象")
        normalized: dict[str, str] = {}
        for key, item in value.items():
            if not isinstance(key, str):
                raise MailboxStoreError("headers 的键必须是字符串")
            if item is None:
                continue
            normalized[key.strip()] = str(item).strip()
        return normalized

    @staticmethod
    def _parse_reference_header(value: str) -> list[str]:
        if not value:
            return []
        return [item for item in value.replace("\r", " ").replace("\n", " ").split(" ") if item]

    @staticmethod
    def _jsonable(value: Any) -> Any:
        if is_dataclass(value):
            return {key: MailboxStore._jsonable(item) for key, item in asdict(value).items()}
        if isinstance(value, list):
            return [MailboxStore._jsonable(item) for item in value]
        if isinstance(value, tuple):
            return [MailboxStore._jsonable(item) for item in value]
        if isinstance(value, dict):
            return {str(key): MailboxStore._jsonable(item) for key, item in value.items()}
        return value

    @staticmethod
    def _json_dumps(value: Any) -> str:
        return json.dumps(value, ensure_ascii=False, sort_keys=True)

    @staticmethod
    def _json_loads(value: Any, default: Any) -> Any:
        if not value:
            return default
        try:
            return json.loads(str(value))
        except json.JSONDecodeError:
            return default

    @staticmethod
    def _build_document_id(mailbox_id: int, method: str, message_id: str) -> str:
        return f"{mailbox_id}|{method}|{message_id}"

    @staticmethod
    def _build_fts_query(query: str) -> str:
        tokens = [item.replace('"', "").strip() for item in query.split() if item.strip()]
        if not tokens:
            return '""'
        return " AND ".join(f'"{token}"' for token in tokens)

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
    def _row_to_folder_entry(row: sqlite3.Row) -> FolderCacheEntry:
        return FolderCacheEntry(
            mailbox_id=int(row["mailbox_id"]),
            method=str(row["method"]),
            folder_id=str(row["folder_id"]),
            name=str(row["name"]),
            display_name=str(row["display_name"]),
            kind=str(row["kind"]),
            total=int(row["total"] or 0),
            unread=int(row["unread"] or 0),
            is_default=bool(row["is_default"]),
            cached_at=str(row["cached_at"]),
            updated_at=str(row["updated_at"]),
        )

    def _row_to_attachment(self, row: sqlite3.Row) -> CachedAttachment:
        return CachedAttachment(
            mailbox_id=int(row["mailbox_id"]),
            method=str(row["method"]),
            message_id=str(row["provider_message_id"]),
            attachment_id=str(row["attachment_id"]),
            name=str(row["name"]),
            content_type=str(row["content_type"] or ""),
            size=int(row["size"] or 0),
            is_inline=bool(row["is_inline"]),
            content_base64=str(row["content_base64"] or ""),
            cached_at=str(row["cached_at"]),
            updated_at=str(row["updated_at"]),
        )

    def _row_to_message_meta(self, row: sqlite3.Row) -> MessageMeta:
        return MessageMeta(
            mailbox_id=int(row["mailbox_id"]),
            method=str(row["method"]),
            message_id=str(row["provider_message_id"]),
            tags=list(self._json_loads(row["tags_json"], [])),
            follow_up=str(row["follow_up"] or ""),
            notes=str(row["notes"] or ""),
            snoozed_until=str(row["snoozed_until"] or ""),
            status=str(row["status"] or row["meta_status"] or "active"),
            updated_at=str(row["updated_at"] or row["meta_updated_at"] or ""),
        )

    def _row_to_cached_message(
        self,
        row: sqlite3.Row,
        *,
        attachments: list[CachedAttachment],
    ) -> CachedMessage:
        meta = None
        if "tags_json" in row.keys():
            meta = MessageMeta(
                mailbox_id=int(row["mailbox_id"]),
                method=str(row["method"]),
                message_id=str(row["provider_message_id"]),
                tags=list(self._json_loads(row["tags_json"], [])),
                follow_up=str(row["follow_up"] or ""),
                notes=str(row["notes"] or ""),
                snoozed_until=str(row["snoozed_until"] or ""),
                status=str(row["meta_status"] or "active"),
                updated_at=str(row["meta_updated_at"] or ""),
            )
        return CachedMessage(
            mailbox_id=int(row["mailbox_id"]),
            mailbox_label=str(row["mailbox_label"] or ""),
            mailbox_email=str(row["mailbox_email"] or ""),
            method=str(row["method"]),
            message_id=str(row["provider_message_id"]),
            subject=str(row["subject"] or ""),
            sender=str(row["sender"] or ""),
            sender_name=str(row["sender_name"] or ""),
            received_at=str(row["received_at"] or ""),
            is_read=bool(row["is_read"]),
            is_flagged=bool(row["is_flagged"]),
            importance=str(row["importance"] or "normal"),
            has_attachments=bool(row["has_attachments"]),
            preview=str(row["preview"] or ""),
            body_text=str(row["body_text"] or ""),
            body_html=str(row["body_html"] or "") or None,
            folder=str(row["folder_id"] or ""),
            internet_message_id=str(row["internet_message_id"] or ""),
            conversation_id=str(row["conversation_id"] or ""),
            to_recipients=list(self._json_loads(row["to_recipients_json"], [])),
            cc_recipients=list(self._json_loads(row["cc_recipients_json"], [])),
            bcc_recipients=list(self._json_loads(row["bcc_recipients_json"], [])),
            headers=dict(self._json_loads(row["headers_json"], {})),
            in_reply_to=str(row["in_reply_to"] or ""),
            references=list(self._json_loads(row["references_json"], [])),
            attachments=attachments,
            meta=meta,
            cached_at=str(row["cached_at"]),
            updated_at=str(row["updated_at"]),
        )

    def _row_to_search_result(self, row: sqlite3.Row) -> MessageSearchResult:
        meta = MessageMeta(
            mailbox_id=int(row["mailbox_id"]),
            method=str(row["method"]),
            message_id=str(row["provider_message_id"]),
            tags=list(self._json_loads(row["tags_json"], [])),
            follow_up=str(row["follow_up"] or ""),
            notes=str(row["notes"] or ""),
            snoozed_until=str(row["snoozed_until"] or ""),
            status=str(row["meta_status"] or "active"),
            updated_at=str(row["meta_updated_at"] or ""),
        )
        return MessageSearchResult(
            mailbox_id=int(row["mailbox_id"]),
            mailbox_label=str(row["mailbox_label"] or ""),
            mailbox_email=str(row["mailbox_email"] or ""),
            method=str(row["method"]),
            message_id=str(row["provider_message_id"]),
            subject=str(row["subject"] or ""),
            sender=str(row["sender"] or ""),
            sender_name=str(row["sender_name"] or ""),
            received_at=str(row["received_at"] or ""),
            is_read=bool(row["is_read"]),
            is_flagged=bool(row["is_flagged"]),
            importance=str(row["importance"] or "normal"),
            has_attachments=bool(row["has_attachments"]),
            preview=str(row["preview"] or ""),
            folder=str(row["folder_id"] or ""),
            internet_message_id=str(row["internet_message_id"] or ""),
            conversation_id=str(row["conversation_id"] or ""),
            meta=meta,
        )

    def _row_to_rule(self, row: sqlite3.Row) -> SavedRule:
        return SavedRule(
            id=int(row["id"]),
            mailbox_id=int(row["mailbox_id"]),
            name=str(row["name"]),
            enabled=bool(row["enabled"]),
            priority=int(row["priority"] or 0),
            conditions=dict(self._json_loads(row["conditions_json"], {})),
            actions=dict(self._json_loads(row["actions_json"], {})),
            created_at=str(row["created_at"]),
            updated_at=str(row["updated_at"]),
        )

    def _row_to_audit_log(self, row: sqlite3.Row) -> AuditLogEntry:
        return AuditLogEntry(
            id=int(row["id"]),
            mailbox_id=int(row["mailbox_id"]) if row["mailbox_id"] is not None else None,
            mailbox_label=str(row["mailbox_label"] or ""),
            mailbox_email=str(row["mailbox_email"] or ""),
            actor=str(row["actor"]),
            action=str(row["action"]),
            target_type=str(row["target_type"]),
            target_id=str(row["target_id"] or ""),
            status=str(row["status"] or "success"),
            details=dict(self._json_loads(row["details_json"], {})),
            created_at=str(row["created_at"]),
        )

    def _row_to_sync_job(self, row: sqlite3.Row) -> SyncJob:
        return SyncJob(
            id=int(row["id"]),
            mailbox_id=int(row["mailbox_id"]),
            method=str(row["method"]),
            requested_by=str(row["requested_by"]),
            scope=dict(self._json_loads(row["scope_json"], {})),
            status=str(row["status"]),
            processed_messages=int(row["processed_messages"] or 0),
            cached_messages=int(row["cached_messages"] or 0),
            folders_synced=int(row["folders_synced"] or 0),
            error=str(row["error"] or ""),
            started_at=str(row["started_at"]),
            finished_at=str(row["finished_at"] or ""),
        )

    @staticmethod
    def _row_to_sync_state(row: sqlite3.Row) -> SyncState:
        return SyncState(
            mailbox_id=int(row["mailbox_id"]),
            method=str(row["method"]),
            folder_id=str(row["folder_id"]),
            last_synced_at=str(row["last_synced_at"]),
            last_message_at=str(row["last_message_at"] or ""),
            cached_messages=int(row["cached_messages"] or 0),
            status=str(row["status"] or "idle"),
            error=str(row["error"] or ""),
            updated_at=str(row["updated_at"]),
        )

    @staticmethod
    def _utc_now() -> str:
        return datetime.now(UTC).isoformat(timespec="seconds").replace("+00:00", "Z")
