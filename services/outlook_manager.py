from __future__ import annotations

import email
import imaplib
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field, replace
from datetime import UTC, datetime, timedelta
from email.header import decode_header
from email.message import Message
from email.utils import getaddresses, parseaddr, parsedate_to_datetime
from html import unescape
from threading import Lock
from typing import Any

import requests

TOKEN_URL_LIVE = "https://login.live.com/oauth20_token.srf"
TOKEN_URL_GRAPH = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
TOKEN_URL_IMAP = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"

GRAPH_MESSAGES_URL = "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
GRAPH_MESSAGE_URL = "https://graph.microsoft.com/v1.0/me/messages/{message_id}"
GRAPH_PROFILE_URL = "https://graph.microsoft.com/v1.0/me"

IMAP_SERVER_OLD = "outlook.office365.com"
IMAP_SERVER_NEW = "outlook.live.com"
IMAP_PORT = 993

METHOD_IMAP_OLD = "imap_old"
METHOD_IMAP_NEW = "imap_new"
METHOD_GRAPH = "graph_api"

FLAGS_PATTERN = re.compile(rb"FLAGS \((?P<flags>[^)]*)\)")
TAG_PATTERN = re.compile(r"<[^>]+>")
IMAP_SUMMARY_FETCH_QUERY = "(FLAGS BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE MESSAGE-ID)])"
TOKEN_CACHE_DEFAULT_TTL_SECONDS = 300


@dataclass(slots=True)
class MailboxConfig:
    email: str
    client_id: str
    refresh_token: str
    proxy: str | None = None
    default_method: str = METHOD_GRAPH
    top: int = 10
    unread_only: bool = False
    keyword: str | None = None


class MailboxError(RuntimeError):
    """邮箱接入过程中的可预期异常。"""

    def __init__(
        self,
        message: str,
        *,
        code: str = "mailbox_error",
        status_code: int = 400,
    ) -> None:
        super().__init__(message)
        self.message = message
        self.code = code
        self.status_code = status_code


@dataclass(slots=True)
class TokenCacheEntry:
    access_token: str
    expires_at: datetime


class OutlookMailboxManager:
    METHOD_LABELS = {
        METHOD_IMAP_OLD: "旧版 IMAP",
        METHOD_IMAP_NEW: "新版 IMAP",
        METHOD_GRAPH: "Graph API",
    }
    METHODS = (METHOD_IMAP_OLD, METHOD_IMAP_NEW, METHOD_GRAPH)

    def __init__(self) -> None:
        self._token_cache: dict[tuple[str, str, str, str], TokenCacheEntry] = {}
        self._token_cache_lock = Lock()

    def get_overview(self, config: MailboxConfig) -> dict[str, Any]:
        cards: list[dict[str, Any]] = []

        with ThreadPoolExecutor(max_workers=len(self.METHODS)) as executor:
            future_map = {
                executor.submit(
                    self.list_messages,
                    replace(config, top=min(max(config.top, 1), 5)),
                    method,
                ): method
                for method in self.METHODS
            }

            for future in as_completed(future_map):
                method = future_map[future]
                card = {
                    "method": method,
                    "label": self.METHOD_LABELS[method],
                    "status": "error",
                    "count": 0,
                    "latest_subject": "",
                    "error": "",
                }
                try:
                    messages = future.result()
                    card["status"] = "ready"
                    card["count"] = len(messages)
                    if messages:
                        card["latest_subject"] = messages[0]["subject"]
                except Exception as exc:  # pragma: no cover - 真实网络分支
                    card["error"] = str(exc)
                cards.append(card)

        cards.sort(key=lambda item: self.METHODS.index(item["method"]))
        return {
            "default_method": config.default_method,
            "generated_at": datetime.now(UTC).isoformat(timespec="seconds").replace("+00:00", "Z"),
            "cards": cards,
        }

    def list_messages(
        self,
        config: MailboxConfig,
        method: str | None = None,
    ) -> list[dict[str, Any]]:
        target_method = method or config.default_method
        if target_method == METHOD_IMAP_OLD:
            return self._list_messages_imap(config, METHOD_IMAP_OLD)
        if target_method == METHOD_IMAP_NEW:
            return self._list_messages_imap(config, METHOD_IMAP_NEW)
        if target_method == METHOD_GRAPH:
            return self._list_messages_graph(config)
        raise MailboxError(f"不支持的方法: {target_method}")

    def resolve_mailbox_email(
        self,
        client_id: str,
        refresh_token: str,
        proxy: str | None = None,
    ) -> str:
        token = self._get_access_token_graph(client_id, refresh_token, proxy)
        response = requests.get(
            GRAPH_PROFILE_URL,
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "mail,userPrincipalName"},
            proxies=self._build_requests_proxies(proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 解析邮箱账号失败")
        payload = response.json()
        email_address = (payload.get("mail") or payload.get("userPrincipalName") or "").strip()
        if not email_address:
            raise MailboxError("无法从授权信息中解析邮箱账号")
        return email_address

    def get_message_detail(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
    ) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._get_message_detail_graph(config, message_id)
        if method in {METHOD_IMAP_OLD, METHOD_IMAP_NEW}:
            return self._get_message_detail_imap(config, method, message_id)
        raise MailboxError(f"不支持的方法: {method}")

    def set_read_state(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
        is_read: bool,
    ) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._set_read_state_graph(config, message_id, is_read)
        if method in {METHOD_IMAP_OLD, METHOD_IMAP_NEW}:
            return self._set_read_state_imap(config, method, message_id, is_read)
        raise MailboxError(f"不支持的方法: {method}")

    def _list_messages_graph(self, config: MailboxConfig) -> list[dict[str, Any]]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        window = max(config.top * 4, 25)
        params: dict[str, Any] = {
            "$top": window,
            "$select": (
                "id,subject,from,receivedDateTime,isRead,"
                "hasAttachments,bodyPreview,internetMessageId"
            ),
            "$orderby": "receivedDateTime desc",
        }
        if config.unread_only:
            params["$filter"] = "isRead eq false"

        response = requests.get(
            GRAPH_MESSAGES_URL,
            headers={
                "Authorization": f"Bearer {token}",
                "Prefer": "outlook.body-content-type='text'",
            },
            params=params,
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 读取邮件失败")

        messages = [self._normalize_graph_summary(item) for item in response.json().get("value", [])]
        return self._filter_messages(messages, config.keyword, config.top)

    def _get_message_detail_graph(self, config: MailboxConfig, message_id: str) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        response = requests.get(
            GRAPH_MESSAGE_URL.format(message_id=message_id),
            headers={"Authorization": f"Bearer {token}"},
            params={
                "$select": (
                    "id,subject,from,toRecipients,ccRecipients,bccRecipients,"
                    "receivedDateTime,isRead,hasAttachments,body,bodyPreview,"
                    "internetMessageId"
                )
            },
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 读取邮件详情失败")
        item = response.json()
        body_text = self._normalize_text(item.get("body", {}).get("content", ""))
        detail = self._normalize_graph_summary(item)
        detail.update(
            {
                "body": body_text,
                "body_text": body_text,
                "to": self._normalize_graph_recipients(item.get("toRecipients", [])),
                "cc": self._normalize_graph_recipients(item.get("ccRecipients", [])),
                "bcc": self._normalize_graph_recipients(item.get("bccRecipients", [])),
                "internet_message_id": item.get("internetMessageId", ""),
            }
        )
        return detail

    def _set_read_state_graph(
        self,
        config: MailboxConfig,
        message_id: str,
        is_read: bool,
    ) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        response = requests.patch(
            GRAPH_MESSAGE_URL.format(message_id=message_id),
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={"isRead": is_read},
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 更新已读状态失败")
        return {
            "message_id": message_id,
            "method": METHOD_GRAPH,
            "is_read": is_read,
            "status": "updated",
        }

    def _list_messages_imap(self, config: MailboxConfig, method: str) -> list[dict[str, Any]]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)

        try:
            status, _ = connection.select("INBOX")
            if status != "OK":
                raise MailboxError("无法选择收件箱")

            search_criteria = ["UNSEEN"] if config.unread_only else ["ALL"]
            status, data = connection.uid("SEARCH", None, *search_criteria)
            if status != "OK" or not data or not data[0]:
                return []

            uid_list = [item for item in data[0].split() if item]
            window = max(config.top * 4, 25)
            target_uids = list(reversed(uid_list[-window:]))

            messages = [self._build_imap_summary(connection, uid, method) for uid in target_uids]
            messages = [item for item in messages if item]
            return self._filter_messages(messages, config.keyword, config.top)
        finally:
            try:
                connection.logout()
            except Exception:  # pragma: no cover - 容错回收
                pass

    def _get_message_detail_imap(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
    ) -> dict[str, Any]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)

        try:
            status, _ = connection.select("INBOX")
            if status != "OK":
                raise MailboxError("无法选择收件箱")

            status, data = connection.uid("FETCH", message_id, "(FLAGS BODY.PEEK[])")
            if status != "OK" or not data:
                raise MailboxError("读取邮件详情失败")

            raw_message = self._extract_raw_message(data)
            flags = self._extract_flags(data)
            message = email.message_from_bytes(raw_message)
            body_text = self._extract_text_body(message)
            detail = self._normalize_imap_message(message_id, message, flags, method)
            detail.update(
                {
                    "body": body_text,
                    "body_text": body_text,
                    "to": self._split_addresses(message.get_all("To", [])),
                    "cc": self._split_addresses(message.get_all("Cc", [])),
                    "bcc": self._split_addresses(message.get_all("Bcc", [])),
                    "internet_message_id": self._decode_header_value(message.get("Message-ID", "")),
                }
            )
            return detail
        finally:
            try:
                connection.logout()
            except Exception:  # pragma: no cover - 容错回收
                pass

    def _set_read_state_imap(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
        is_read: bool,
    ) -> dict[str, Any]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)

        try:
            status, _ = connection.select("INBOX")
            if status != "OK":
                raise MailboxError("无法选择收件箱")

            operation = "+FLAGS" if is_read else "-FLAGS"
            status, _ = connection.uid("STORE", message_id, operation, "(\\Seen)")
            if status != "OK":
                raise MailboxError("更新已读状态失败")

            return {
                "message_id": message_id,
                "method": method,
                "is_read": is_read,
                "status": "updated",
            }
        finally:
            try:
                connection.logout()
            except Exception:  # pragma: no cover - 容错回收
                pass

    def _build_imap_summary(
        self,
        connection: imaplib.IMAP4_SSL,
        uid: bytes,
        method: str,
    ) -> dict[str, Any] | None:
        # 列表刷新只需要摘要信息，避免把整封邮件正文都拉回来解析。
        status, data = connection.uid("FETCH", uid, IMAP_SUMMARY_FETCH_QUERY)
        if status != "OK" or not data:
            return None

        raw_message = self._extract_raw_message(data)
        flags = self._extract_flags(data)
        message = email.message_from_bytes(raw_message)
        return self._normalize_imap_message(uid.decode("utf-8"), message, flags, method)

    def _normalize_imap_message(
        self,
        message_id: str,
        message: Message,
        flags: set[str],
        method: str,
    ) -> dict[str, Any]:
        sender_name, sender_address = parseaddr(self._decode_header_value(message.get("From", "")))
        preview = self._summarize_text(self._extract_text_body(message))
        received_at = self._normalize_received_at(message.get("Date", ""))
        return {
            "id": message_id,
            "message_id": message_id,
            "method": method,
            "source": self.METHOD_LABELS[method],
            "subject": self._decode_header_value(message.get("Subject", "无主题")) or "无主题",
            "sender": sender_address or sender_name or "未知发件人",
            "sender_name": sender_name,
            "received_at": received_at,
            "is_read": "\\Seen" in flags,
            "has_attachments": any(part.get_filename() for part in message.walk()),
            "preview": preview,
            "internet_message_id": self._decode_header_value(message.get("Message-ID", "")),
        }

    def _normalize_graph_summary(self, item: dict[str, Any]) -> dict[str, Any]:
        email_address = item.get("from", {}).get("emailAddress", {})
        message_id = item.get("id", "")
        return {
            "id": message_id,
            "message_id": message_id,
            "method": METHOD_GRAPH,
            "source": self.METHOD_LABELS[METHOD_GRAPH],
            "subject": item.get("subject") or "无主题",
            "sender": email_address.get("address", "未知发件人"),
            "sender_name": email_address.get("name", ""),
            "received_at": item.get("receivedDateTime", ""),
            "is_read": bool(item.get("isRead")),
            "has_attachments": bool(item.get("hasAttachments")),
            "preview": self._normalize_text(item.get("bodyPreview", "")),
            "internet_message_id": item.get("internetMessageId", ""),
        }

    def _get_access_token_for_imap(self, config: MailboxConfig, method: str) -> str:
        if method == METHOD_IMAP_OLD:
            return self._get_access_token_old(config.client_id, config.refresh_token, config.proxy)
        return self._get_access_token_imap(config.client_id, config.refresh_token, config.proxy)

    def _get_access_token_old(self, client_id: str, refresh_token: str, proxy: str | None) -> str:
        return self._fetch_access_token(
            cache_scope="live",
            client_id=client_id,
            refresh_token=refresh_token,
            proxy=proxy,
            url=TOKEN_URL_LIVE,
            payload={
                "client_id": client_id,
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
            },
            missing_token_message="旧版 IMAP access_token 缺失",
            request_error_message="旧版 IMAP 获取 access_token 失败",
        )

    def _get_access_token_imap(self, client_id: str, refresh_token: str, proxy: str | None) -> str:
        return self._fetch_access_token(
            cache_scope="imap",
            client_id=client_id,
            refresh_token=refresh_token,
            proxy=proxy,
            url=TOKEN_URL_IMAP,
            payload={
                "client_id": client_id,
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
                "scope": "https://outlook.office.com/IMAP.AccessAsUser.All offline_access",
            },
            missing_token_message="新版 IMAP access_token 缺失",
            request_error_message="新版 IMAP 获取 access_token 失败",
        )

    def _get_access_token_graph(self, client_id: str, refresh_token: str, proxy: str | None) -> str:
        return self._fetch_access_token(
            cache_scope="graph",
            client_id=client_id,
            refresh_token=refresh_token,
            proxy=proxy,
            url=TOKEN_URL_GRAPH,
            payload={
                "client_id": client_id,
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
                "scope": "https://graph.microsoft.com/.default",
            },
            missing_token_message="Graph API access_token 缺失",
            request_error_message="Graph API 获取 access_token 失败",
        )

    def _fetch_access_token(
        self,
        *,
        cache_scope: str,
        client_id: str,
        refresh_token: str,
        proxy: str | None,
        url: str,
        payload: dict[str, str],
        missing_token_message: str,
        request_error_message: str,
    ) -> str:
        cache_key = self._build_token_cache_key(cache_scope, client_id, refresh_token, proxy)
        cached_token = self._get_cached_access_token(cache_key)
        if cached_token:
            return cached_token

        response = requests.post(
            url,
            data=payload,
            proxies=self._build_requests_proxies(proxy),
            timeout=30,
        )
        self._raise_for_response(response, request_error_message)
        response_payload = response.json()
        token = response_payload.get("access_token")
        if not token:
            raise MailboxError(missing_token_message)
        self._store_cached_access_token(cache_key, token, response_payload.get("expires_in"))
        return token

    @staticmethod
    def _build_token_cache_key(
        cache_scope: str,
        client_id: str,
        refresh_token: str,
        proxy: str | None,
    ) -> tuple[str, str, str, str]:
        return (cache_scope, client_id, refresh_token, proxy or "")

    def _get_cached_access_token(self, cache_key: tuple[str, str, str, str]) -> str | None:
        now = datetime.now(UTC)
        with self._token_cache_lock:
            entry = self._token_cache.get(cache_key)
            if not entry:
                return None
            if entry.expires_at <= now:
                self._token_cache.pop(cache_key, None)
                return None
            return entry.access_token

    def _store_cached_access_token(
        self,
        cache_key: tuple[str, str, str, str],
        token: str,
        expires_in: Any,
    ) -> None:
        ttl_seconds = self._normalize_token_ttl(expires_in)
        with self._token_cache_lock:
            self._token_cache[cache_key] = TokenCacheEntry(
                access_token=token,
                expires_at=datetime.now(UTC) + timedelta(seconds=ttl_seconds),
            )

    @staticmethod
    def _normalize_token_ttl(expires_in: Any) -> int:
        try:
            raw_ttl = int(expires_in)
        except (TypeError, ValueError):
            raw_ttl = TOKEN_CACHE_DEFAULT_TTL_SECONDS

        raw_ttl = max(raw_ttl, 10)
        safety_buffer = min(60, max(raw_ttl // 10, 5))
        return max(raw_ttl - safety_buffer, 5)

    def _open_imap_connection(self, account: str, access_token: str, method: str) -> imaplib.IMAP4_SSL:
        server = IMAP_SERVER_OLD if method == METHOD_IMAP_OLD else IMAP_SERVER_NEW
        try:
            connection = imaplib.IMAP4_SSL(server, IMAP_PORT)
            auth_string = f"user={account}\1auth=Bearer {access_token}\1\1".encode("utf-8")
            connection.authenticate("XOAUTH2", lambda _: auth_string)
            return connection
        except Exception as exc:  # pragma: no cover - 真实网络分支
            raise MailboxError(f"连接 {self.METHOD_LABELS[method]} 失败: {exc}") from exc

    @staticmethod
    def _build_requests_proxies(proxy: str | None) -> dict[str, str] | None:
        if not proxy:
            return None
        normalized = proxy if proxy.startswith("http://") or proxy.startswith("https://") else f"http://{proxy}"
        return {"http": normalized, "https": normalized}

    @staticmethod
    def _raise_for_response(response: requests.Response, message: str) -> None:
        if response.status_code < 400:
            return

        body = response.text[:300]
        if "service abuse mode" in body:
            raise MailboxError(f"{message}: 账号疑似进入风控模式")
        raise MailboxError(f"{message}: {response.status_code} {body}")

    @staticmethod
    def _extract_raw_message(data: list[Any]) -> bytes:
        for item in data:
            if isinstance(item, tuple) and len(item) >= 2 and isinstance(item[1], bytes):
                return item[1]
        raise MailboxError("邮件原文为空")

    @staticmethod
    def _extract_flags(data: list[Any]) -> set[str]:
        for item in data:
            if isinstance(item, tuple):
                header = item[0]
                if isinstance(header, bytes):
                    match = FLAGS_PATTERN.search(header)
                    if match:
                        return {part.decode("utf-8", "ignore") for part in match.group("flags").split()}
        return set()

    @staticmethod
    def _normalize_received_at(value: str) -> str:
        try:
            return parsedate_to_datetime(value).isoformat()
        except Exception:
            return value

    @staticmethod
    def _normalize_graph_recipients(recipients: list[dict[str, Any]]) -> list[str]:
        values: list[str] = []
        for item in recipients:
            email_address = item.get("emailAddress", {})
            name = email_address.get("name")
            address = email_address.get("address")
            if name and address:
                values.append(f"{name} <{address}>")
            elif address:
                values.append(address)
        return values

    @staticmethod
    def _split_addresses(values: list[str]) -> list[str]:
        addresses: list[str] = []
        for _, address in getaddresses(values):
            if address:
                addresses.append(address)
        return addresses

    @staticmethod
    def _summarize_text(text: str, limit: int = 180) -> str:
        cleaned = " ".join(text.split())
        if len(cleaned) <= limit:
            return cleaned
        return cleaned[: limit - 1].rstrip() + "…"

    @staticmethod
    def _normalize_text(value: str) -> str:
        text = TAG_PATTERN.sub(" ", value)
        return " ".join(unescape(text).split())

    def _filter_messages(
        self,
        messages: list[dict[str, Any]],
        keyword: str | None,
        top: int,
    ) -> list[dict[str, Any]]:
        if keyword:
            needle = keyword.casefold()
            messages = [
                item
                for item in messages
                if needle in f"{item['subject']} {item['sender']} {item['preview']}".casefold()
            ]
        return messages[:top]

    def _decode_header_value(self, header_value: Any) -> str:
        if not header_value:
            return ""

        fragments: list[str] = []
        for part, charset in decode_header(str(header_value)):
            if isinstance(part, bytes):
                encoding = charset or "utf-8"
                try:
                    fragments.append(part.decode(encoding, "replace"))
                except LookupError:
                    fragments.append(part.decode("utf-8", "replace"))
            else:
                fragments.append(str(part))
        return "".join(fragments).strip()

    def _extract_text_body(self, message: Message) -> str:
        if message.is_multipart():
            html_fallback = ""
            for part in message.walk():
                content_type = part.get_content_type()
                if part.get_content_disposition() == "attachment":
                    continue

                payload = part.get_payload(decode=True)
                if payload is None:
                    continue

                charset = part.get_content_charset() or "utf-8"
                try:
                    text = payload.decode(charset, "replace")
                except LookupError:
                    text = payload.decode("utf-8", "replace")

                if content_type == "text/plain":
                    return self._normalize_text(text)
                if content_type == "text/html" and not html_fallback:
                    html_fallback = self._normalize_text(text)
            return html_fallback

        payload = message.get_payload(decode=True)
        if payload is None:
            raw = message.get_payload()
            return self._normalize_text(raw if isinstance(raw, str) else "")

        charset = message.get_content_charset() or "utf-8"
        try:
            return self._normalize_text(payload.decode(charset, "replace"))
        except LookupError:
            return self._normalize_text(payload.decode("utf-8", "replace"))


@dataclass(slots=True)
class MailboxQuery:
    method: str
    top: int = 10
    unread_only: bool = False
    keyword: str = ""
    folder: str = "INBOX"


@dataclass(slots=True)
class MessageSummary:
    method: str
    message_id: str
    subject: str
    sender: str
    sender_name: str
    received_at: str
    is_read: bool
    has_attachments: bool
    preview: str
    source: str
    folder: str = "INBOX"
    internet_message_id: str = ""


@dataclass(slots=True)
class MessageDetailRequest:
    method: str
    message_id: str
    folder: str = "INBOX"


@dataclass(slots=True)
class ReadStateUpdateRequest:
    method: str
    message_id: str
    is_read: bool
    folder: str = "INBOX"


@dataclass(slots=True)
class MessageDetail(MessageSummary):
    body_text: str = ""
    body_html: str | None = None
    to_recipients: list[str] = field(default_factory=list)
    cc_recipients: list[str] = field(default_factory=list)
    bcc_recipients: list[str] = field(default_factory=list)
    headers: dict[str, str] = field(default_factory=dict)
    conversation_id: str = ""


@dataclass(slots=True)
class MessageListResult:
    method: str
    total: int
    returned: int
    messages: list[MessageSummary] = field(default_factory=list)
    folder: str = "INBOX"

    def __len__(self) -> int:
        return len(self.messages)


@dataclass(slots=True)
class MethodOverview:
    method: str
    label: str
    healthy: bool
    status: str
    message: str
    message_count: int = 0
    latest_subject: str = ""
    latest_sender: str = ""


@dataclass(slots=True)
class MailboxOverview:
    email: str
    checked_at: str
    methods: list[MethodOverview] = field(default_factory=list)


@dataclass(slots=True)
class ReadStateUpdateResult:
    method: str
    message_id: str
    is_read: bool
    status: str
    source: str


MailMethodOverview = MethodOverview
MailboxDetailRequest = MessageDetailRequest
MessageListRequest = MailboxQuery
UpdateReadStateRequest = ReadStateUpdateRequest
UpdateReadStateResult = ReadStateUpdateResult
MailboxServiceError = MailboxError


class MailboxManager:
    """对外暴露的统一服务接口，内部复用已有的 OutlookMailboxManager。"""

    def __init__(self, inner: OutlookMailboxManager | None = None) -> None:
        self._inner = inner or OutlookMailboxManager()

    def get_overview(self, config: MailboxConfig) -> MailboxOverview:
        raw = self._inner.get_overview(config)
        methods = [
            MethodOverview(
                method=item.get("method", ""),
                label=item.get("label", item.get("method", "")),
                healthy=item.get("status") == "ready",
                status=item.get("status", "error"),
                message=item.get("error") or ("连接正常" if item.get("status") == "ready" else "连接失败"),
                message_count=int(item.get("count", 0) or 0),
                latest_subject=item.get("latest_subject", ""),
                latest_sender=item.get("latest_sender", ""),
            )
            for item in raw.get("cards", [])
        ]
        return MailboxOverview(
            email=config.email,
            checked_at=raw.get("generated_at", datetime.now(UTC).isoformat(timespec="seconds") + "Z"),
            methods=methods,
        )

    def list_messages(self, config: MailboxConfig, request: MailboxQuery) -> MessageListResult:
        effective = replace(
            config,
            default_method=request.method,
            top=request.top,
            unread_only=request.unread_only,
            keyword=request.keyword or None,
        )
        raw_messages = self._inner.list_messages(effective, request.method)
        messages = [self._to_summary(item, request.method, request.folder) for item in raw_messages]
        return MessageListResult(
            method=request.method,
            total=len(raw_messages),
            returned=len(messages),
            messages=messages,
            folder=request.folder,
        )

    def get_message_detail(self, config: MailboxConfig, request: MessageDetailRequest) -> MessageDetail:
        raw = self._inner.get_message_detail(config, request.method, request.message_id)
        return self._to_detail(raw, request.method, request.folder)

    def update_read_state(self, config: MailboxConfig, request: ReadStateUpdateRequest) -> ReadStateUpdateResult:
        raw = self._inner.set_read_state(config, request.method, request.message_id, request.is_read)
        return ReadStateUpdateResult(
            method=raw.get("method", request.method),
            message_id=raw.get("message_id", request.message_id),
            is_read=bool(raw.get("is_read", request.is_read)),
            status=raw.get("status", "updated"),
            source=OutlookMailboxManager.METHOD_LABELS.get(request.method, request.method),
        )

    def resolve_mailbox_email(
        self,
        *,
        client_id: str,
        refresh_token: str,
        proxy: str | None = None,
    ) -> str:
        return self._inner.resolve_mailbox_email(client_id, refresh_token, proxy)

    def _to_summary(self, item: dict[str, Any], method: str, folder: str) -> MessageSummary:
        return MessageSummary(
            method=item.get("method", method),
            message_id=item.get("id") or item.get("message_id", ""),
            subject=item.get("subject", "无主题"),
            sender=item.get("sender", "未知发件人"),
            sender_name=item.get("sender_name", ""),
            received_at=item.get("received_at", ""),
            is_read=bool(item.get("is_read", False)),
            has_attachments=bool(item.get("has_attachments", False)),
            preview=item.get("preview", ""),
            source=OutlookMailboxManager.METHOD_LABELS.get(item.get("method", method), method),
            folder=folder,
            internet_message_id=item.get("internet_message_id", ""),
        )

    def _to_detail(self, item: dict[str, Any], method: str, folder: str) -> MessageDetail:
        summary = self._to_summary(item, method, folder)
        return MessageDetail(
            method=summary.method,
            message_id=summary.message_id,
            subject=summary.subject,
            sender=summary.sender,
            sender_name=summary.sender_name,
            received_at=summary.received_at,
            is_read=summary.is_read,
            has_attachments=summary.has_attachments,
            preview=summary.preview,
            source=summary.source,
            folder=summary.folder,
            internet_message_id=summary.internet_message_id,
            body_text=item.get("body", ""),
            body_html=item.get("body_html"),
            to_recipients=list(item.get("to", [])),
            cc_recipients=list(item.get("cc", [])),
            bcc_recipients=list(item.get("bcc", [])),
            headers=dict(item.get("headers", {})),
            conversation_id=item.get("conversation_id", ""),
        )
