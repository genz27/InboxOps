from __future__ import annotations

import base64
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
from urllib.parse import quote

import requests

TOKEN_URL_LIVE = "https://login.live.com/oauth20_token.srf"
TOKEN_URL_GRAPH = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
TOKEN_URL_IMAP = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"

GRAPH_MESSAGES_URL = "https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages"
GRAPH_MESSAGE_URL = "https://graph.microsoft.com/v1.0/me/messages/{message_id}"
GRAPH_CREATE_MESSAGE_URL = "https://graph.microsoft.com/v1.0/me/messages"
GRAPH_SEND_MAIL_URL = "https://graph.microsoft.com/v1.0/me/sendMail"
GRAPH_MESSAGE_SEND_URL = "https://graph.microsoft.com/v1.0/me/messages/{message_id}/send"
GRAPH_MESSAGE_MOVE_URL = "https://graph.microsoft.com/v1.0/me/messages/{message_id}/move"
GRAPH_MESSAGE_ATTACHMENTS_URL = "https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments"
GRAPH_MESSAGE_ATTACHMENT_URL = "https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments/{attachment_id}"
GRAPH_MESSAGE_CREATE_REPLY_URL = "https://graph.microsoft.com/v1.0/me/messages/{message_id}/createReply"
GRAPH_MESSAGE_CREATE_REPLY_ALL_URL = "https://graph.microsoft.com/v1.0/me/messages/{message_id}/createReplyAll"
GRAPH_MESSAGE_CREATE_FORWARD_URL = "https://graph.microsoft.com/v1.0/me/messages/{message_id}/createForward"
GRAPH_MAIL_FOLDERS_URL = "https://graph.microsoft.com/v1.0/me/mailFolders"
GRAPH_MAIL_FOLDER_URL = "https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}"
GRAPH_MAIL_FOLDER_CHILDREN_URL = "https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/childFolders"
GRAPH_PROFILE_URL = "https://graph.microsoft.com/v1.0/me"

IMAP_SERVER_OLD = "outlook.office365.com"
IMAP_SERVER_NEW = "outlook.live.com"
IMAP_PORT = 993

METHOD_IMAP_OLD = "imap_old"
METHOD_IMAP_NEW = "imap_new"
METHOD_GRAPH = "graph_api"

DEFAULT_FOLDER = "INBOX"
DEFAULT_READ_STATE = "all"
DEFAULT_IMPORTANCE = "all"
DEFAULT_SORT_ORDER = "desc"
TOKEN_CACHE_DEFAULT_TTL_SECONDS = 300

FLAGS_PATTERN = re.compile(rb"FLAGS \((?P<flags>[^)]*)\)")
IMAP_LIST_PATTERN = re.compile(rb'\((?P<flags>[^)]*)\)\s+"(?P<delimiter>[^"]*)"\s+(?P<name>.+)')
TAG_PATTERN = re.compile(r"<[^>]+>")
SCRIPT_STYLE_PATTERN = re.compile(r"(?is)<(script|style).*?>.*?</\1>")
HTML_BREAK_PATTERN = re.compile(
    r"(?is)</?(?:br|p|div|section|article|header|footer|li|tr|table|blockquote|h[1-6]|hr)[^>]*>"
)
IMAP_SUMMARY_FETCH_QUERY = "(FLAGS BODY.PEEK[])"

GRAPH_WELL_KNOWN_FOLDERS = {
    "inbox": {"label": "收件箱", "kind": "inbox"},
    "archive": {"label": "归档", "kind": "archive"},
    "deleteditems": {"label": "已删除", "kind": "trash"},
    "drafts": {"label": "草稿箱", "kind": "drafts"},
    "junkemail": {"label": "垃圾邮件", "kind": "junk"},
    "sentitems": {"label": "已发送", "kind": "sent"},
}

FOLDER_KIND_ALIASES = {
    "inbox": "inbox",
    "archive": "archive",
    "deleteditems": "trash",
    "deleted": "trash",
    "trash": "trash",
    "drafts": "drafts",
    "junkemail": "junk",
    "junk": "junk",
    "spam": "junk",
    "sentitems": "sent",
    "sent": "sent",
}

IMAP_SPECIAL_USE_ALIASES = {
    "\\inbox": "inbox",
    "\\archive": "archive",
    "\\trash": "trash",
    "\\deleted": "trash",
    "\\drafts": "drafts",
    "\\junk": "junk",
    "\\spam": "junk",
    "\\sent": "sent",
    "\\sentmail": "sent",
}


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
    folder: str = DEFAULT_FOLDER
    page: int = 1
    page_size: int = 20
    read_state: str = DEFAULT_READ_STATE
    has_attachments_only: bool = False
    flagged_only: bool = False
    importance: str = DEFAULT_IMPORTANCE
    sort_order: str = DEFAULT_SORT_ORDER


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


@dataclass(slots=True)
class MailboxFolder:
    id: str
    name: str
    display_name: str
    total: int = 0
    unread: int = 0
    is_default: bool = False
    kind: str = "custom"


@dataclass(slots=True)
class AttachmentSummary:
    id: str
    name: str
    content_type: str
    size: int
    is_inline: bool = False


@dataclass(slots=True)
class MailboxQuery:
    method: str
    top: int = 10
    unread_only: bool = False
    keyword: str = ""
    folder: str = DEFAULT_FOLDER
    page: int = 1
    page_size: int = 20
    read_state: str = DEFAULT_READ_STATE
    has_attachments_only: bool = False
    flagged_only: bool = False
    importance: str = DEFAULT_IMPORTANCE
    sort_order: str = DEFAULT_SORT_ORDER


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
    folder: str = DEFAULT_FOLDER
    internet_message_id: str = ""
    is_flagged: bool = False
    importance: str = "normal"
    conversation_id: str = ""


@dataclass(slots=True)
class MessageDetailRequest:
    method: str
    message_id: str
    folder: str = DEFAULT_FOLDER


@dataclass(slots=True)
class ReadStateUpdateRequest:
    method: str
    message_id: str
    is_read: bool
    folder: str = DEFAULT_FOLDER


@dataclass(slots=True)
class FlagStateUpdateRequest:
    method: str
    message_id: str
    is_flagged: bool
    folder: str = DEFAULT_FOLDER


@dataclass(slots=True)
class MessageMoveRequest:
    method: str
    message_id: str
    destination_folder: str
    folder: str = DEFAULT_FOLDER


@dataclass(slots=True)
class MessageDeleteRequest:
    method: str
    message_id: str
    folder: str = DEFAULT_FOLDER


@dataclass(slots=True)
class ComposeAttachment:
    name: str
    content_base64: str
    content_type: str = "application/octet-stream"
    is_inline: bool = False


@dataclass(slots=True)
class DraftSaveRequest:
    method: str
    subject: str
    body_text: str = ""
    body_html: str | None = None
    to_recipients: list[str] = field(default_factory=list)
    cc_recipients: list[str] = field(default_factory=list)
    bcc_recipients: list[str] = field(default_factory=list)
    attachments: list[ComposeAttachment] = field(default_factory=list)
    importance: str = "normal"
    draft_id: str = ""


@dataclass(slots=True)
class MessageSendRequest:
    method: str
    subject: str
    body_text: str = ""
    body_html: str | None = None
    to_recipients: list[str] = field(default_factory=list)
    cc_recipients: list[str] = field(default_factory=list)
    bcc_recipients: list[str] = field(default_factory=list)
    attachments: list[ComposeAttachment] = field(default_factory=list)
    importance: str = "normal"
    save_to_sent_items: bool = True


@dataclass(slots=True)
class MessageReplyRequest:
    method: str
    message_id: str
    folder: str = DEFAULT_FOLDER
    body_text: str = ""
    body_html: str | None = None
    comment: str = ""
    attachments: list[ComposeAttachment] = field(default_factory=list)
    importance: str = "normal"


@dataclass(slots=True)
class MessageForwardRequest:
    method: str
    message_id: str
    to_recipients: list[str]
    folder: str = DEFAULT_FOLDER
    body_text: str = ""
    body_html: str | None = None
    comment: str = ""
    attachments: list[ComposeAttachment] = field(default_factory=list)
    importance: str = "normal"


@dataclass(slots=True)
class AttachmentDownloadRequest:
    method: str
    message_id: str
    attachment_id: str
    folder: str = DEFAULT_FOLDER


@dataclass(slots=True)
class AttachmentUploadRequest:
    method: str
    message_id: str
    attachments: list[ComposeAttachment] = field(default_factory=list)
    folder: str = "drafts"


@dataclass(slots=True)
class FolderCreateRequest:
    method: str
    display_name: str
    parent_folder: str = ""


@dataclass(slots=True)
class FolderRenameRequest:
    method: str
    folder_id: str
    display_name: str


@dataclass(slots=True)
class FolderDeleteRequest:
    method: str
    folder_id: str


@dataclass(slots=True)
class MessageDetail(MessageSummary):
    body_text: str = ""
    body_html: str | None = None
    to_recipients: list[str] = field(default_factory=list)
    cc_recipients: list[str] = field(default_factory=list)
    bcc_recipients: list[str] = field(default_factory=list)
    headers: dict[str, str] = field(default_factory=dict)
    attachments: list[AttachmentSummary] = field(default_factory=list)
    conversation_id: str = ""


@dataclass(slots=True)
class MessageListResult:
    method: str
    total: int
    returned: int
    messages: list[MessageSummary] = field(default_factory=list)
    folder: str = DEFAULT_FOLDER
    page: int = 1
    page_size: int = 20
    total_pages: int = 0
    has_prev: bool = False
    has_next: bool = False

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


@dataclass(slots=True)
class FlagStateUpdateResult:
    method: str
    message_id: str
    is_flagged: bool
    status: str
    source: str


@dataclass(slots=True)
class MessageMoveResult:
    method: str
    message_id: str
    source_folder: str
    destination_folder: str
    status: str
    source: str


@dataclass(slots=True)
class MessageDeleteResult:
    method: str
    message_id: str
    folder: str
    status: str
    source: str


@dataclass(slots=True)
class MessageSendResult:
    method: str
    message_id: str
    status: str
    folder: str
    source: str


@dataclass(slots=True)
class AttachmentDownloadResult:
    method: str
    message_id: str
    attachment_id: str
    name: str
    content_type: str
    size: int
    content_base64: str
    source: str


@dataclass(slots=True)
class AttachmentUploadResult:
    method: str
    message_id: str
    attachments: list[AttachmentSummary] = field(default_factory=list)
    status: str = "uploaded"
    source: str = ""


@dataclass(slots=True)
class FolderMutationResult:
    method: str
    folder_id: str
    display_name: str
    status: str
    source: str


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
                    replace(
                        config,
                        top=min(max(config.top, 1), 5),
                        page=1,
                        page_size=min(max(config.top, 1), 5),
                    ),
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
                    "latest_sender": "",
                    "error": "",
                }
                try:
                    messages = future.result()
                    card["status"] = "ready"
                    card["count"] = len(messages)
                    if messages:
                        card["latest_subject"] = messages[0]["subject"]
                        card["latest_sender"] = messages[0]["sender"]
                except Exception as exc:  # pragma: no cover - 真实网络分支
                    card["error"] = str(exc)
                cards.append(card)

        cards.sort(key=lambda item: self.METHODS.index(item["method"]))
        return {
            "default_method": config.default_method,
            "generated_at": datetime.now(UTC).isoformat(timespec="seconds").replace("+00:00", "Z"),
            "cards": cards,
        }

    def list_folders(self, config: MailboxConfig, method: str | None = None) -> list[dict[str, Any]]:
        target_method = method or config.default_method
        if target_method == METHOD_GRAPH:
            return self._list_folders_graph(config)
        if target_method in {METHOD_IMAP_OLD, METHOD_IMAP_NEW}:
            return self._list_folders_imap(config, target_method)
        raise MailboxError(f"不支持的方法: {target_method}")

    def list_messages(self, config: MailboxConfig, method: str | None = None) -> list[dict[str, Any]]:
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

    def get_message_detail(self, config: MailboxConfig, method: str, message_id: str) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._get_message_detail_graph(config, message_id)
        if method in {METHOD_IMAP_OLD, METHOD_IMAP_NEW}:
            return self._get_message_detail_imap(config, method, message_id)
        raise MailboxError(f"不支持的方法: {method}")

    def set_read_state(self, config: MailboxConfig, method: str, message_id: str, is_read: bool) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._set_read_state_graph(config, message_id, is_read)
        if method in {METHOD_IMAP_OLD, METHOD_IMAP_NEW}:
            return self._set_read_state_imap(config, method, message_id, is_read)
        raise MailboxError(f"不支持的方法: {method}")

    def set_flag_state(self, config: MailboxConfig, method: str, message_id: str, is_flagged: bool) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._set_flag_state_graph(config, message_id, is_flagged)
        if method in {METHOD_IMAP_OLD, METHOD_IMAP_NEW}:
            return self._set_flag_state_imap(config, method, message_id, is_flagged)
        raise MailboxError(f"不支持的方法: {method}")

    def move_message(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
        destination_folder: str,
    ) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._move_message_graph(config, message_id, destination_folder)
        if method in {METHOD_IMAP_OLD, METHOD_IMAP_NEW}:
            return self._move_message_imap(config, method, message_id, destination_folder)
        raise MailboxError(f"不支持的方法: {method}")

    def delete_message(self, config: MailboxConfig, method: str, message_id: str) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._delete_message_graph(config, message_id)
        if method in {METHOD_IMAP_OLD, METHOD_IMAP_NEW}:
            return self._delete_message_imap(config, method, message_id)
        raise MailboxError(f"不支持的方法: {method}")

    def save_draft(self, config: MailboxConfig, method: str, payload: dict[str, Any]) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._save_draft_graph(config, payload)
        raise MailboxError("当前接入方式暂不支持草稿保存", code="method_not_supported", status_code=501)

    def send_message(self, config: MailboxConfig, method: str, payload: dict[str, Any]) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._send_message_graph(config, payload)
        raise MailboxError("当前接入方式暂不支持发信", code="method_not_supported", status_code=501)

    def reply_message(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
        payload: dict[str, Any],
        *,
        reply_all: bool = False,
    ) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._reply_message_graph(config, message_id, payload, reply_all=reply_all)
        raise MailboxError("当前接入方式暂不支持回复邮件", code="method_not_supported", status_code=501)

    def forward_message(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
        payload: dict[str, Any],
    ) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._forward_message_graph(config, message_id, payload)
        raise MailboxError("当前接入方式暂不支持转发邮件", code="method_not_supported", status_code=501)

    def upload_attachment(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
        payload: dict[str, Any],
    ) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._upload_attachment_graph(config, message_id, payload)
        raise MailboxError("当前接入方式暂不支持附件上传", code="method_not_supported", status_code=501)

    def download_attachment(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
        attachment_id: str,
    ) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._download_attachment_graph(config, message_id, attachment_id)
        if method in {METHOD_IMAP_OLD, METHOD_IMAP_NEW}:
            return self._download_attachment_imap(config, method, message_id, attachment_id)
        raise MailboxError(f"不支持的方法: {method}")

    def create_folder(self, config: MailboxConfig, method: str, payload: dict[str, Any]) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._create_folder_graph(config, payload)
        if method in {METHOD_IMAP_OLD, METHOD_IMAP_NEW}:
            return self._create_folder_imap(config, method, payload)
        raise MailboxError(f"不支持的方法: {method}")

    def rename_folder(self, config: MailboxConfig, method: str, payload: dict[str, Any]) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._rename_folder_graph(config, payload)
        if method in {METHOD_IMAP_OLD, METHOD_IMAP_NEW}:
            return self._rename_folder_imap(config, method, payload)
        raise MailboxError(f"不支持的方法: {method}")

    def delete_folder(self, config: MailboxConfig, method: str, payload: dict[str, Any]) -> dict[str, Any]:
        if method == METHOD_GRAPH:
            return self._delete_folder_graph(config, payload)
        if method in {METHOD_IMAP_OLD, METHOD_IMAP_NEW}:
            return self._delete_folder_imap(config, method, payload)
        raise MailboxError(f"不支持的方法: {method}")

    def _list_folders_graph(self, config: MailboxConfig) -> list[dict[str, Any]]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        response = requests.get(
            GRAPH_MAIL_FOLDERS_URL,
            headers={"Authorization": f"Bearer {token}"},
            params={"$top": 200, "$select": "id,displayName,totalItemCount,unreadItemCount"},
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 读取文件夹失败")
        folders = [self._normalize_graph_folder(item) for item in response.json().get("value", [])]
        return self._ensure_default_graph_folders(folders)

    def _list_messages_graph(self, config: MailboxConfig) -> list[dict[str, Any]]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        window = min(max(max(config.page * config.page_size, config.top) * 4, 50), 500)
        filters = self._build_graph_filters(config)
        params: dict[str, Any] = {
            "$top": window,
            "$select": (
                "id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview,"
                "internetMessageId,flag,importance,conversationId"
            ),
            "$orderby": f"receivedDateTime {self._normalize_sort_order(config.sort_order)}",
        }
        if filters:
            params["$filter"] = " and ".join(filters)

        response = requests.get(
            GRAPH_MESSAGES_URL.format(folder_id=quote(self._normalize_graph_folder_id(config.folder), safe="")),
            headers={
                "Authorization": f"Bearer {token}",
                "Prefer": "outlook.body-content-type='text'",
            },
            params=params,
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 读取邮件失败")

        messages = [self._normalize_graph_summary(item, folder=config.folder) for item in response.json().get("value", [])]
        return self._filter_messages(messages, config)

    def _get_message_detail_graph(self, config: MailboxConfig, message_id: str) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        response = requests.get(
            GRAPH_MESSAGE_URL.format(message_id=message_id),
            headers={"Authorization": f"Bearer {token}"},
            params={
                "$select": (
                    "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,"
                    "isRead,hasAttachments,body,bodyPreview,internetMessageId,flag,"
                    "importance,conversationId,internetMessageHeaders"
                ),
                "$expand": "attachments($select=id,name,contentType,size,isInline)",
            },
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 读取邮件详情失败")
        item = response.json()
        body_payload = item.get("body", {}) if isinstance(item.get("body"), dict) else {}
        body_raw = str(body_payload.get("content", "") or "")
        body_type = str(body_payload.get("contentType", "") or "").casefold()
        body_html = self._sanitize_html(body_raw) if body_type == "html" else None
        body_text = self._normalize_text(body_raw) if body_raw else self._normalize_text(item.get("bodyPreview", ""))
        detail = self._normalize_graph_summary(item, folder=config.folder)
        detail.update(
            {
                "body": body_text,
                "body_text": body_text,
                "body_html": body_html,
                "to": self._normalize_graph_recipients(item.get("toRecipients", [])),
                "cc": self._normalize_graph_recipients(item.get("ccRecipients", [])),
                "bcc": self._normalize_graph_recipients(item.get("bccRecipients", [])),
                "headers": self._normalize_graph_headers(item.get("internetMessageHeaders", [])),
                "internet_message_id": item.get("internetMessageId", ""),
                "conversation_id": item.get("conversationId", ""),
                "attachments": [
                    self._normalize_graph_attachment(attachment)
                    for attachment in item.get("attachments", [])
                    if isinstance(attachment, dict)
                ],
            }
        )
        return detail

    def _set_read_state_graph(self, config: MailboxConfig, message_id: str, is_read: bool) -> dict[str, Any]:
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

    def _set_flag_state_graph(self, config: MailboxConfig, message_id: str, is_flagged: bool) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        response = requests.patch(
            GRAPH_MESSAGE_URL.format(message_id=message_id),
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={"flag": {"flagStatus": "flagged" if is_flagged else "notFlagged"}},
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 更新星标状态失败")
        return {
            "message_id": message_id,
            "method": METHOD_GRAPH,
            "is_flagged": is_flagged,
            "status": "updated",
        }

    def _move_message_graph(
        self,
        config: MailboxConfig,
        message_id: str,
        destination_folder: str,
    ) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        resolved_destination = self._resolve_graph_destination_id(config, destination_folder)
        response = requests.post(
            GRAPH_MESSAGE_MOVE_URL.format(message_id=message_id),
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={"destinationId": resolved_destination},
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 移动邮件失败")
        return {
            "message_id": message_id,
            "method": METHOD_GRAPH,
            "source_folder": config.folder,
            "destination_folder": destination_folder,
            "status": "moved",
        }

    def _delete_message_graph(self, config: MailboxConfig, message_id: str) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        response = requests.delete(
            GRAPH_MESSAGE_URL.format(message_id=message_id),
            headers={"Authorization": f"Bearer {token}"},
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 删除邮件失败")
        return {
            "message_id": message_id,
            "method": METHOD_GRAPH,
            "folder": config.folder,
            "status": "deleted",
        }

    def _save_draft_graph(self, config: MailboxConfig, payload: dict[str, Any]) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        draft_id = self._normalize_optional_payload_text(payload.get("draft_message_id") or payload.get("message_id"))
        message_payload = self._build_graph_message_payload(payload)
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        if draft_id:
            response = requests.patch(
                GRAPH_MESSAGE_URL.format(message_id=draft_id),
                headers=headers,
                json=message_payload,
                proxies=self._build_requests_proxies(config.proxy),
                timeout=30,
            )
            self._raise_for_response(response, "Graph API 更新草稿失败")
        else:
            response = requests.post(
                GRAPH_CREATE_MESSAGE_URL,
                headers=headers,
                json=message_payload,
                proxies=self._build_requests_proxies(config.proxy),
                timeout=30,
            )
            self._raise_for_response(response, "Graph API 保存草稿失败")
            body = response.json() if response.content else {}
            draft_id = str(body.get("id", "") or "")

        if not draft_id:
            raise MailboxError("Graph API 未返回草稿标识")

        attachments = payload.get("attachments")
        if isinstance(attachments, list):
            for attachment in attachments:
                if not isinstance(attachment, dict):
                    continue
                self._upload_attachment_graph(config, draft_id, attachment)

        detail = self._get_message_detail_graph(replace(config, folder="drafts"), draft_id)
        detail.update(
            {
                "method": METHOD_GRAPH,
                "status": "draft_saved",
                "folder": "drafts",
            }
        )
        return detail

    def _send_message_graph(self, config: MailboxConfig, payload: dict[str, Any]) -> dict[str, Any]:
        draft_id = self._normalize_optional_payload_text(payload.get("draft_message_id") or payload.get("message_id"))
        draft = self._save_draft_graph(config, payload) if not draft_id or self._payload_has_compose_updates(payload) else None
        resolved_draft_id = draft.get("message_id", "") if isinstance(draft, dict) else draft_id or ""
        if not resolved_draft_id:
            raise MailboxError("缺少待发送草稿标识", code="invalid_message_id")
        result = self._send_existing_graph_draft(config, resolved_draft_id)
        if draft:
            result.update(
                {
                    "subject": draft.get("subject", result.get("subject", "")),
                    "sender": draft.get("sender", result.get("sender", "")),
                    "sender_name": draft.get("sender_name", result.get("sender_name", "")),
                    "to": draft.get("to", []),
                    "cc": draft.get("cc", []),
                    "bcc": draft.get("bcc", []),
                    "attachments": draft.get("attachments", []),
                    "conversation_id": draft.get("conversation_id", result.get("conversation_id", "")),
                }
            )
        return result

    def _reply_message_graph(
        self,
        config: MailboxConfig,
        message_id: str,
        payload: dict[str, Any],
        *,
        reply_all: bool = False,
    ) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        create_url = GRAPH_MESSAGE_CREATE_REPLY_ALL_URL if reply_all else GRAPH_MESSAGE_CREATE_REPLY_URL
        response = requests.post(
            create_url.format(message_id=message_id),
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={},
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 创建回复草稿失败")
        body = response.json() if response.content else {}
        draft_id = str(body.get("id", "") or "")
        if not draft_id:
            raise MailboxError("Graph API 未返回回复草稿标识")

        next_payload = dict(payload)
        next_payload["draft_message_id"] = draft_id
        saved = self._save_draft_graph(config, next_payload)
        if not self._payload_send_now(payload, default=True):
            saved["status"] = "draft_saved"
            return saved
        result = self._send_existing_graph_draft(config, draft_id)
        result.update(
            {
                "subject": saved.get("subject", result.get("subject", "")),
                "attachments": saved.get("attachments", []),
                "conversation_id": saved.get("conversation_id", result.get("conversation_id", "")),
                "reply_all": reply_all,
            }
        )
        return result

    def _forward_message_graph(self, config: MailboxConfig, message_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        response = requests.post(
            GRAPH_MESSAGE_CREATE_FORWARD_URL.format(message_id=message_id),
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={},
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 创建转发草稿失败")
        body = response.json() if response.content else {}
        draft_id = str(body.get("id", "") or "")
        if not draft_id:
            raise MailboxError("Graph API 未返回转发草稿标识")

        next_payload = dict(payload)
        next_payload["draft_message_id"] = draft_id
        saved = self._save_draft_graph(config, next_payload)
        if not self._payload_send_now(payload, default=True):
            saved["status"] = "draft_saved"
            return saved
        result = self._send_existing_graph_draft(config, draft_id)
        result.update(
            {
                "subject": saved.get("subject", result.get("subject", "")),
                "attachments": saved.get("attachments", []),
                "conversation_id": saved.get("conversation_id", result.get("conversation_id", "")),
            }
        )
        return result

    def _send_existing_graph_draft(self, config: MailboxConfig, draft_id: str) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        detail = self._get_message_detail_graph(replace(config, folder="drafts"), draft_id)
        response = requests.post(
            GRAPH_MESSAGE_SEND_URL.format(message_id=draft_id),
            headers={"Authorization": f"Bearer {token}"},
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 发送邮件失败")
        detail.update(
            {
                "method": METHOD_GRAPH,
                "message_id": draft_id,
                "status": "sent",
                "folder": "sentitems",
            }
        )
        return detail

    def _upload_attachment_graph(self, config: MailboxConfig, message_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        name = self._require_payload_text(payload, "name", "缺少附件名称")
        content_base64 = self._require_payload_text(payload, "content_base64", "缺少附件内容")
        response = requests.post(
            GRAPH_MESSAGE_ATTACHMENTS_URL.format(message_id=message_id),
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": name,
                "contentType": self._normalize_optional_payload_text(payload.get("content_type")) or "application/octet-stream",
                "contentBytes": content_base64,
                "isInline": bool(payload.get("is_inline", False)),
                "contentId": self._normalize_optional_payload_text(payload.get("content_id")) or "",
            },
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 上传附件失败")
        attachment = response.json() if response.content else {}
        normalized = self._normalize_graph_attachment(attachment)
        normalized.update(
            {
                "content_base64": content_base64,
                "content_id": self._normalize_optional_payload_text(attachment.get("contentId"))
                or self._normalize_optional_payload_text(payload.get("content_id"))
                or "",
            }
        )
        return normalized

    def _download_attachment_graph(self, config: MailboxConfig, message_id: str, attachment_id: str) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        response = requests.get(
            GRAPH_MESSAGE_ATTACHMENT_URL.format(message_id=message_id, attachment_id=attachment_id),
            headers={"Authorization": f"Bearer {token}"},
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 下载附件失败")
        attachment = response.json()
        normalized = self._normalize_graph_attachment(attachment)
        normalized.update(
            {
                "content_base64": str(attachment.get("contentBytes", "") or ""),
                "content_id": str(attachment.get("contentId", "") or ""),
            }
        )
        return normalized

    def _create_folder_graph(self, config: MailboxConfig, payload: dict[str, Any]) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        display_name = self._require_payload_text(payload, "display_name", "缺少文件夹名称")
        parent_folder = self._normalize_optional_payload_text(payload.get("parent_folder_id") or payload.get("parent_folder"))
        url = (
            GRAPH_MAIL_FOLDER_CHILDREN_URL.format(folder_id=quote(self._normalize_graph_folder_id(parent_folder), safe=""))
            if parent_folder
            else GRAPH_MAIL_FOLDERS_URL
        )
        response = requests.post(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={"displayName": display_name},
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 新建文件夹失败")
        folder = self._normalize_graph_folder(response.json())
        folder["status"] = "created"
        folder["parent_folder_id"] = parent_folder or ""
        return folder

    def _rename_folder_graph(self, config: MailboxConfig, payload: dict[str, Any]) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        folder_id = self._require_payload_text(payload, "folder_id", "缺少文件夹标识")
        display_name = self._require_payload_text(payload, "display_name", "缺少新的文件夹名称")
        response = requests.patch(
            GRAPH_MAIL_FOLDER_URL.format(folder_id=quote(self._normalize_graph_folder_id(folder_id), safe="")),
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={"displayName": display_name},
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 重命名文件夹失败")
        folder = self._normalize_graph_folder(response.json()) if response.content else {
            "id": folder_id,
            "name": display_name,
            "display_name": display_name,
            "kind": self._resolve_folder_kind(display_name),
            "total": 0,
            "unread": 0,
            "is_default": False,
        }
        folder["status"] = "renamed"
        return folder

    def _delete_folder_graph(self, config: MailboxConfig, payload: dict[str, Any]) -> dict[str, Any]:
        token = self._get_access_token_graph(config.client_id, config.refresh_token, config.proxy)
        folder_id = self._require_payload_text(payload, "folder_id", "缺少文件夹标识")
        response = requests.delete(
            GRAPH_MAIL_FOLDER_URL.format(folder_id=quote(self._normalize_graph_folder_id(folder_id), safe="")),
            headers={"Authorization": f"Bearer {token}"},
            proxies=self._build_requests_proxies(config.proxy),
            timeout=30,
        )
        self._raise_for_response(response, "Graph API 删除文件夹失败")
        return {
            "id": folder_id,
            "status": "deleted",
        }

    def _list_folders_imap(self, config: MailboxConfig, method: str) -> list[dict[str, Any]]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)
        try:
            return self._list_imap_folders_from_connection(connection)
        finally:
            self._logout_imap(connection)

    def _list_messages_imap(self, config: MailboxConfig, method: str) -> list[dict[str, Any]]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)
        folder_name = self._normalize_imap_folder_name(config.folder)

        try:
            self._select_imap_folder(connection, folder_name)
            search_criteria = self._build_imap_search_criteria(config)
            status, data = connection.uid("SEARCH", None, *search_criteria)
            if status != "OK" or not data or not data[0]:
                return []

            uid_list = [item for item in data[0].split() if item]
            window = min(max(max(config.page * config.page_size, config.top) * 4, 50), 500)
            target_uids = list(reversed(uid_list[-window:]))
            messages = [self._build_imap_summary(connection, uid, method, folder_name) for uid in target_uids]
            messages = [item for item in messages if item]
            return self._filter_messages(messages, config)
        finally:
            self._logout_imap(connection)

    def _get_message_detail_imap(self, config: MailboxConfig, method: str, message_id: str) -> dict[str, Any]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)
        folder_name = self._normalize_imap_folder_name(config.folder)

        try:
            self._select_imap_folder(connection, folder_name)
            status, data = connection.uid("FETCH", message_id, "(FLAGS BODY.PEEK[])")
            if status != "OK" or not data:
                raise MailboxError("读取邮件详情失败")

            raw_message = self._extract_raw_message(data)
            flags = self._extract_flags(data)
            message = email.message_from_bytes(raw_message)
            body_text, body_html = self._extract_message_bodies(message)
            detail = self._normalize_imap_message(message_id, message, flags, method, folder_name)
            detail.update(
                {
                    "body": body_text,
                    "body_text": body_text,
                    "body_html": body_html,
                    "to": self._split_addresses(message.get_all("To", [])),
                    "cc": self._split_addresses(message.get_all("Cc", [])),
                    "bcc": self._split_addresses(message.get_all("Bcc", [])),
                    "headers": self._normalize_message_headers(message),
                    "internet_message_id": self._decode_header_value(message.get("Message-ID", "")),
                    "conversation_id": self._decode_header_value(message.get("Thread-Index", "")),
                    "attachments": self._extract_imap_attachments(message),
                }
            )
            return detail
        finally:
            self._logout_imap(connection)

    def _set_read_state_imap(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
        is_read: bool,
    ) -> dict[str, Any]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)
        folder_name = self._normalize_imap_folder_name(config.folder)

        try:
            self._select_imap_folder(connection, folder_name)
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
            self._logout_imap(connection)

    def _set_flag_state_imap(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
        is_flagged: bool,
    ) -> dict[str, Any]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)
        folder_name = self._normalize_imap_folder_name(config.folder)

        try:
            self._select_imap_folder(connection, folder_name)
            operation = "+FLAGS" if is_flagged else "-FLAGS"
            status, _ = connection.uid("STORE", message_id, operation, "(\\Flagged)")
            if status != "OK":
                raise MailboxError("更新星标状态失败")
            return {
                "message_id": message_id,
                "method": method,
                "is_flagged": is_flagged,
                "status": "updated",
            }
        finally:
            self._logout_imap(connection)

    def _move_message_imap(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
        destination_folder: str,
    ) -> dict[str, Any]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)
        folder_name = self._normalize_imap_folder_name(config.folder)

        try:
            self._select_imap_folder(connection, folder_name)
            target_folder = self._resolve_imap_destination_folder(connection, destination_folder)
            status, _ = connection.uid("COPY", message_id, self._quote_imap_mailbox(target_folder))
            if status != "OK":
                raise MailboxError("移动邮件失败")
            status, _ = connection.uid("STORE", message_id, "+FLAGS", "(\\Deleted)")
            if status != "OK":
                raise MailboxError("移动邮件失败")
            expunge_status, _ = connection.expunge()
            if expunge_status != "OK":
                raise MailboxError("移动邮件后执行 expunge 失败")
            return {
                "message_id": message_id,
                "method": method,
                "source_folder": folder_name,
                "destination_folder": target_folder,
                "status": "moved",
            }
        finally:
            self._logout_imap(connection)

    def _delete_message_imap(self, config: MailboxConfig, method: str, message_id: str) -> dict[str, Any]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)
        folder_name = self._normalize_imap_folder_name(config.folder)

        try:
            self._select_imap_folder(connection, folder_name)
            trash_folder = self._resolve_imap_destination_folder(connection, "trash", allow_missing=True)
            if trash_folder and trash_folder.casefold() != folder_name.casefold():
                copy_status, _ = connection.uid("COPY", message_id, self._quote_imap_mailbox(trash_folder))
                if copy_status == "OK":
                    delete_status, _ = connection.uid("STORE", message_id, "+FLAGS", "(\\Deleted)")
                    if delete_status != "OK":
                        raise MailboxError("删除邮件失败")
                    connection.expunge()
                    return {
                        "message_id": message_id,
                        "method": method,
                        "folder": folder_name,
                        "status": "moved_to_trash",
                    }

            status, _ = connection.uid("STORE", message_id, "+FLAGS", "(\\Deleted)")
            if status != "OK":
                raise MailboxError("删除邮件失败")
            connection.expunge()
            return {
                "message_id": message_id,
                "method": method,
                "folder": folder_name,
                "status": "deleted",
            }
        finally:
            self._logout_imap(connection)

    def _download_attachment_imap(
        self,
        config: MailboxConfig,
        method: str,
        message_id: str,
        attachment_id: str,
    ) -> dict[str, Any]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)
        folder_name = self._normalize_imap_folder_name(config.folder)

        try:
            self._select_imap_folder(connection, folder_name)
            status, data = connection.uid("FETCH", message_id, "(BODY.PEEK[])")
            if status != "OK" or not data:
                raise MailboxError("下载附件失败")
            raw_message = self._extract_raw_message(data)
            message = email.message_from_bytes(raw_message)
            attachment = self._extract_imap_attachment_content(message, attachment_id)
            if not attachment:
                raise MailboxError("附件不存在", code="attachment_not_found", status_code=404)
            return attachment
        finally:
            self._logout_imap(connection)

    def _create_folder_imap(self, config: MailboxConfig, method: str, payload: dict[str, Any]) -> dict[str, Any]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)
        display_name = self._require_payload_text(payload, "display_name", "缺少文件夹名称")
        parent_folder = self._normalize_optional_payload_text(payload.get("parent_folder_id") or payload.get("parent_folder"))
        folder_name = f"{parent_folder}/{display_name}" if parent_folder else display_name
        try:
            status, _ = connection.create(self._quote_imap_mailbox(folder_name))
            if status != "OK":
                raise MailboxError("新建文件夹失败")
            return {
                "id": folder_name,
                "name": folder_name,
                "display_name": display_name,
                "kind": self._resolve_folder_kind(display_name),
                "total": 0,
                "unread": 0,
                "is_default": False,
                "status": "created",
            }
        finally:
            self._logout_imap(connection)

    def _rename_folder_imap(self, config: MailboxConfig, method: str, payload: dict[str, Any]) -> dict[str, Any]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)
        folder_id = self._require_payload_text(payload, "folder_id", "缺少文件夹标识")
        display_name = self._require_payload_text(payload, "display_name", "缺少新的文件夹名称")
        parent_folder = self._normalize_optional_payload_text(payload.get("parent_folder_id") or payload.get("parent_folder"))
        new_folder_name = f"{parent_folder}/{display_name}" if parent_folder else display_name
        try:
            status, _ = connection.rename(self._quote_imap_mailbox(folder_id), self._quote_imap_mailbox(new_folder_name))
            if status != "OK":
                raise MailboxError("重命名文件夹失败")
            return {
                "id": new_folder_name,
                "name": new_folder_name,
                "display_name": display_name,
                "kind": self._resolve_folder_kind(display_name),
                "total": 0,
                "unread": 0,
                "is_default": False,
                "status": "renamed",
                "previous_id": folder_id,
            }
        finally:
            self._logout_imap(connection)

    def _delete_folder_imap(self, config: MailboxConfig, method: str, payload: dict[str, Any]) -> dict[str, Any]:
        token = self._get_access_token_for_imap(config, method)
        connection = self._open_imap_connection(config.email, token, method)
        folder_id = self._require_payload_text(payload, "folder_id", "缺少文件夹标识")
        try:
            status, _ = connection.delete(self._quote_imap_mailbox(folder_id))
            if status != "OK":
                raise MailboxError("删除文件夹失败")
            return {
                "id": folder_id,
                "status": "deleted",
            }
        finally:
            self._logout_imap(connection)

    def _build_imap_summary(
        self,
        connection: imaplib.IMAP4_SSL,
        uid: bytes,
        method: str,
        folder_name: str,
    ) -> dict[str, Any] | None:
        status, data = connection.uid("FETCH", uid, IMAP_SUMMARY_FETCH_QUERY)
        if status != "OK" or not data:
            return None
        raw_message = self._extract_raw_message(data)
        flags = self._extract_flags(data)
        message = email.message_from_bytes(raw_message)
        return self._normalize_imap_message(uid.decode("utf-8"), message, flags, method, folder_name)

    def _normalize_imap_message(
        self,
        message_id: str,
        message: Message,
        flags: set[str],
        method: str,
        folder_name: str,
    ) -> dict[str, Any]:
        sender_name, sender_address = parseaddr(self._decode_header_value(message.get("From", "")))
        body_text, _ = self._extract_message_bodies(message)
        return {
            "id": message_id,
            "message_id": message_id,
            "method": method,
            "source": self.METHOD_LABELS[method],
            "subject": self._decode_header_value(message.get("Subject", "无主题")) or "无主题",
            "sender": sender_address or sender_name or "未知发件人",
            "sender_name": sender_name,
            "received_at": self._normalize_received_at(message.get("Date", "")),
            "is_read": "\\Seen" in flags,
            "has_attachments": bool(self._extract_imap_attachments(message)),
            "preview": self._summarize_text(body_text),
            "internet_message_id": self._decode_header_value(message.get("Message-ID", "")),
            "folder": folder_name,
            "is_flagged": "\\Flagged" in flags,
            "importance": self._extract_importance_from_message(message),
            "conversation_id": self._decode_header_value(message.get("Thread-Index", "")),
        }

    def _normalize_graph_summary(self, item: dict[str, Any], *, folder: str) -> dict[str, Any]:
        email_address = item.get("from", {}).get("emailAddress", {})
        flag = item.get("flag", {}) if isinstance(item.get("flag"), dict) else {}
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
            "folder": folder or DEFAULT_FOLDER,
            "is_flagged": flag.get("flagStatus") == "flagged",
            "importance": self._normalize_importance(item.get("importance")),
            "conversation_id": item.get("conversationId", ""),
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
        response = requests.post(url, data=payload, proxies=self._build_requests_proxies(proxy), timeout=30)
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

    def _store_cached_access_token(self, cache_key: tuple[str, str, str, str], token: str, expires_in: Any) -> None:
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

    def _normalize_graph_folder(self, item: dict[str, Any]) -> dict[str, Any]:
        folder_id = str(item.get("id", "") or "")
        display_name = str(item.get("displayName", "") or folder_id or DEFAULT_FOLDER)
        kind = self._resolve_folder_kind(display_name)
        return {
            "id": folder_id,
            "name": kind if kind != "custom" else display_name,
            "display_name": display_name,
            "total": int(item.get("totalItemCount", 0) or 0),
            "unread": int(item.get("unreadItemCount", 0) or 0),
            "is_default": kind == "inbox",
            "kind": kind,
        }

    def _ensure_default_graph_folders(self, folders: list[dict[str, Any]]) -> list[dict[str, Any]]:
        seen_ids = {str(item.get("id", "")).casefold() for item in folders}
        next_folders = list(folders)
        for folder_id, definition in GRAPH_WELL_KNOWN_FOLDERS.items():
            if folder_id.casefold() in seen_ids:
                continue
            next_folders.append(
                {
                    "id": folder_id,
                    "name": definition["kind"],
                    "display_name": definition["label"],
                    "total": 0,
                    "unread": 0,
                    "is_default": definition["kind"] == "inbox",
                    "kind": definition["kind"],
                }
            )
        return self._sort_folders(next_folders)

    def _list_imap_folders_from_connection(self, connection: imaplib.IMAP4_SSL) -> list[dict[str, Any]]:
        status, data = connection.list()
        if status != "OK":
            raise MailboxError("读取 IMAP 文件夹失败")
        folders = [item for line in data or [] if isinstance(line, bytes) and (item := self._normalize_imap_folder(line))]
        return self._ensure_default_imap_folders(folders)

    def _normalize_imap_folder(self, line: bytes) -> dict[str, Any] | None:
        match = IMAP_LIST_PATTERN.match(line)
        if not match:
            return None
        flags = {item.decode("utf-8", "ignore").casefold() for item in match.group("flags").split()}
        raw_name = match.group("name").decode("utf-8", "replace").strip().strip('"')
        display_name = self._decode_imap_utf7(raw_name)
        kind = self._resolve_folder_kind(display_name, flags=flags)
        return {
            "id": raw_name,
            "name": kind if kind != "custom" else raw_name,
            "display_name": display_name or raw_name,
            "total": 0,
            "unread": 0,
            "is_default": kind == "inbox",
            "kind": kind,
        }

    def _ensure_default_imap_folders(self, folders: list[dict[str, Any]]) -> list[dict[str, Any]]:
        if any(str(item.get("kind", "")).casefold() == "inbox" for item in folders):
            return self._sort_folders(folders)
        return self._sort_folders(
            [
                {
                    "id": DEFAULT_FOLDER,
                    "name": "inbox",
                    "display_name": DEFAULT_FOLDER,
                    "total": 0,
                    "unread": 0,
                    "is_default": True,
                    "kind": "inbox",
                },
                *folders,
            ]
        )

    def _resolve_folder_kind(self, value: str, *, flags: set[str] | None = None) -> str:
        if flags:
            for flag in flags:
                if flag in IMAP_SPECIAL_USE_ALIASES:
                    return IMAP_SPECIAL_USE_ALIASES[flag]
        normalized = value.strip().casefold().replace(" ", "")
        return FOLDER_KIND_ALIASES.get(normalized, "custom")

    def _sort_folders(self, folders: list[dict[str, Any]]) -> list[dict[str, Any]]:
        order = {"inbox": 0, "sent": 1, "drafts": 2, "archive": 3, "trash": 4, "junk": 5, "custom": 99}
        return sorted(
            folders,
            key=lambda item: (order.get(str(item.get("kind", "custom")), 99), str(item.get("display_name", "")).casefold()),
        )

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
    def _normalize_graph_headers(headers: list[dict[str, Any]]) -> dict[str, str]:
        normalized: dict[str, str] = {}
        for item in headers:
            if not isinstance(item, dict):
                continue
            name = str(item.get("name", "") or "").strip()
            value = str(item.get("value", "") or "").strip()
            if name:
                normalized[name] = value
        return normalized

    @staticmethod
    def _normalize_graph_attachment(attachment: dict[str, Any]) -> dict[str, Any]:
        return {
            "id": str(attachment.get("id", "") or ""),
            "name": str(attachment.get("name", "") or ""),
            "content_type": str(attachment.get("contentType", "") or ""),
            "size": int(attachment.get("size", 0) or 0),
            "is_inline": bool(attachment.get("isInline", False)),
        }

    def _build_graph_message_payload(self, payload: dict[str, Any]) -> dict[str, Any]:
        body_html = self._normalize_optional_payload_text(payload.get("body_html"))
        body_text = self._normalize_optional_payload_text(payload.get("body_text") or payload.get("body"))
        body_type = "HTML" if body_html else "Text"
        body_content = body_html or body_text or ""
        message: dict[str, Any] = {
            "subject": self._normalize_optional_payload_text(payload.get("subject")) or "",
            "body": {
                "contentType": body_type,
                "content": body_content,
            },
        }

        to_recipients = self._build_graph_recipient_payload(payload.get("to_recipients") or payload.get("to"))
        cc_recipients = self._build_graph_recipient_payload(payload.get("cc_recipients") or payload.get("cc"))
        bcc_recipients = self._build_graph_recipient_payload(payload.get("bcc_recipients") or payload.get("bcc"))
        if to_recipients:
            message["toRecipients"] = to_recipients
        if cc_recipients:
            message["ccRecipients"] = cc_recipients
        if bcc_recipients:
            message["bccRecipients"] = bcc_recipients
        return message

    def _build_graph_recipient_payload(self, values: Any) -> list[dict[str, Any]]:
        recipients: list[dict[str, Any]] = []
        if values in (None, ""):
            return recipients
        if not isinstance(values, list):
            raise MailboxError("收件人列表必须是数组", code="invalid_recipients")
        for item in values:
            if not isinstance(item, str) or not item.strip():
                raise MailboxError("收件人列表中存在无效地址", code="invalid_recipients")
            name, address = parseaddr(item.strip())
            if not address:
                raise MailboxError("收件人地址无效", code="invalid_recipients")
            recipients.append(
                {
                    "emailAddress": {
                        "name": name or address,
                        "address": address,
                    }
                }
            )
        return recipients

    @staticmethod
    def _payload_send_now(payload: dict[str, Any], *, default: bool) -> bool:
        raw = payload.get("send_now")
        return default if raw is None else bool(raw)

    @staticmethod
    def _payload_has_compose_updates(payload: dict[str, Any]) -> bool:
        return any(
            key in payload
            for key in (
                "subject",
                "body",
                "body_text",
                "body_html",
                "to",
                "to_recipients",
                "cc",
                "cc_recipients",
                "bcc",
                "bcc_recipients",
                "attachments",
            )
        )

    @staticmethod
    def _require_payload_text(payload: dict[str, Any], key: str, error_message: str) -> str:
        value = OutlookMailboxManager._normalize_optional_payload_text(payload.get(key))
        if not value:
            raise MailboxError(error_message, code=f"invalid_{key}")
        return value

    @staticmethod
    def _normalize_optional_payload_text(value: Any) -> str | None:
        if value is None:
            return None
        if not isinstance(value, str):
            raise MailboxError("请求参数类型错误", code="invalid_payload")
        cleaned = value.strip()
        return cleaned or None

    @staticmethod
    def _split_addresses(values: list[str]) -> list[str]:
        addresses: list[str] = []
        for name, address in getaddresses(values):
            if address:
                addresses.append(f"{name} <{address}>" if name else address)
        return addresses

    @staticmethod
    def _summarize_text(text: str, limit: int = 180) -> str:
        cleaned = " ".join(text.split())
        if len(cleaned) <= limit:
            return cleaned
        return cleaned[: limit - 1].rstrip() + "…"

    @staticmethod
    def _sanitize_html(value: str) -> str:
        html = SCRIPT_STYLE_PATTERN.sub(" ", value or "")
        html = re.sub(r"(?is)\son[a-z]+\s*=\s*(\"[^\"]*\"|'[^']*')", "", html)
        return re.sub(r"(?i)javascript:", "", html)

    @staticmethod
    def _normalize_text(value: str) -> str:
        text = SCRIPT_STYLE_PATTERN.sub(" ", value or "")
        text = HTML_BREAK_PATTERN.sub("\n", text)
        text = TAG_PATTERN.sub(" ", text)
        normalized_lines = [" ".join(unescape(line).split()) for line in text.splitlines()]
        cleaned_lines = [line for line in normalized_lines if line]
        return "\n".join(cleaned_lines)

    def _filter_messages(self, messages: list[dict[str, Any]], config: MailboxConfig) -> list[dict[str, Any]]:
        filtered = list(messages)
        if config.keyword:
            needle = config.keyword.casefold()
            filtered = [
                item
                for item in filtered
                if needle in f"{item['subject']} {item['sender']} {item['preview']}".casefold()
            ]
        if config.has_attachments_only:
            filtered = [item for item in filtered if bool(item.get("has_attachments"))]
        if config.flagged_only:
            filtered = [item for item in filtered if bool(item.get("is_flagged"))]
        if config.importance != DEFAULT_IMPORTANCE:
            filtered = [
                item
                for item in filtered
                if self._normalize_importance(item.get("importance")) == self._normalize_importance(config.importance)
            ]
        reverse = self._normalize_sort_order(config.sort_order) == "desc"
        filtered.sort(key=lambda item: str(item.get("received_at", "")), reverse=reverse)
        start = max((config.page - 1) * config.page_size, 0)
        end = start + config.page_size
        return filtered[start:end]

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
        return self._extract_message_bodies(message)[0]

    def _extract_message_bodies(self, message: Message) -> tuple[str, str | None]:
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
                    return self._normalize_text(text), html_fallback or None
                if content_type == "text/html" and not html_fallback:
                    html_fallback = self._sanitize_html(text)
            return self._normalize_text(html_fallback), html_fallback or None

        payload = message.get_payload(decode=True)
        if payload is None:
            raw = message.get_payload()
            raw_text = raw if isinstance(raw, str) else ""
            normalized = self._normalize_text(raw_text)
            if message.get_content_type() == "text/html":
                return normalized, self._sanitize_html(raw_text)
            return normalized, None

        charset = message.get_content_charset() or "utf-8"
        try:
            decoded = payload.decode(charset, "replace")
        except LookupError:
            decoded = payload.decode("utf-8", "replace")
        normalized = self._normalize_text(decoded)
        if message.get_content_type() == "text/html":
            return normalized, self._sanitize_html(decoded)
        return normalized, None

    def _extract_imap_attachments(self, message: Message) -> list[dict[str, Any]]:
        attachments: list[dict[str, Any]] = []
        for index, part in enumerate(message.walk(), start=1):
            file_name = part.get_filename()
            disposition = (part.get_content_disposition() or "").casefold()
            if not file_name and disposition != "attachment":
                continue
            payload = part.get_payload(decode=True) or b""
            attachments.append(
                {
                    "id": f"part-{index}",
                    "name": self._decode_header_value(file_name or f"attachment-{index}"),
                    "content_type": part.get_content_type(),
                    "size": len(payload),
                    "is_inline": disposition == "inline",
                }
            )
        return attachments

    def _extract_imap_attachment_content(self, message: Message, attachment_id: str) -> dict[str, Any] | None:
        normalized_attachment_id = str(attachment_id or "").strip()
        for index, part in enumerate(message.walk(), start=1):
            candidate_id = f"part-{index}"
            if candidate_id != normalized_attachment_id:
                continue
            file_name = part.get_filename()
            disposition = (part.get_content_disposition() or "").casefold()
            payload = part.get_payload(decode=True) or b""
            return {
                "id": candidate_id,
                "name": self._decode_header_value(file_name or f"attachment-{index}"),
                "content_type": part.get_content_type(),
                "size": len(payload),
                "is_inline": disposition == "inline",
                "content_base64": base64.b64encode(payload).decode("ascii"),
            }
        return None

    def _normalize_message_headers(self, message: Message) -> dict[str, str]:
        normalized: dict[str, str] = {}
        for key, value in message.items():
            decoded_key = self._decode_header_value(key)
            decoded_value = self._decode_header_value(value)
            if not decoded_key:
                continue
            if decoded_key in normalized:
                normalized[decoded_key] = f"{normalized[decoded_key]} | {decoded_value}"
            else:
                normalized[decoded_key] = decoded_value
        return normalized

    def _extract_importance_from_message(self, message: Message) -> str:
        joined = " ".join(
            value.casefold()
            for value in [
                self._decode_header_value(message.get("Importance", "")),
                self._decode_header_value(message.get("Priority", "")),
                self._decode_header_value(message.get("X-Priority", "")),
            ]
            if value
        )
        if "high" in joined or joined.startswith("1"):
            return "high"
        if "low" in joined or joined.startswith("5"):
            return "low"
        return "normal"

    @staticmethod
    def _normalize_importance(value: Any) -> str:
        normalized = str(value or "").strip().casefold()
        if normalized in {"high", "normal", "low"}:
            return normalized
        if normalized in {"", DEFAULT_IMPORTANCE}:
            return DEFAULT_IMPORTANCE
        return "normal"

    @staticmethod
    def _normalize_sort_order(value: Any) -> str:
        return "asc" if str(value or "").strip().casefold() == "asc" else "desc"

    def _build_graph_filters(self, config: MailboxConfig) -> list[str]:
        filters: list[str] = []
        read_state = config.read_state
        if config.unread_only and read_state == DEFAULT_READ_STATE:
            read_state = "unread"
        if read_state == "unread":
            filters.append("isRead eq false")
        elif read_state == "read":
            filters.append("isRead eq true")
        if config.has_attachments_only:
            filters.append("hasAttachments eq true")
        if config.flagged_only:
            filters.append("flag/flagStatus eq 'flagged'")
        importance = self._normalize_importance(config.importance)
        if importance in {"high", "normal", "low"}:
            filters.append(f"importance eq '{importance}'")
        return filters

    @staticmethod
    def _normalize_graph_folder_id(folder: str) -> str:
        normalized = str(folder or "").strip()
        if not normalized or normalized == DEFAULT_FOLDER:
            return "inbox"
        lowered = normalized.casefold()
        if lowered in GRAPH_WELL_KNOWN_FOLDERS:
            return lowered
        for key, definition in GRAPH_WELL_KNOWN_FOLDERS.items():
            if definition["kind"] == lowered:
                return key
        return normalized

    def _resolve_graph_destination_id(self, config: MailboxConfig, destination_folder: str) -> str:
        target = str(destination_folder or "").strip()
        if not target:
            raise MailboxError("缺少目标文件夹", code="invalid_destination_folder")
        folders = self._list_folders_graph(config)
        target_key = target.casefold()
        for item in folders:
            if target_key in {
                str(item.get("id", "")).casefold(),
                str(item.get("name", "")).casefold(),
                str(item.get("display_name", "")).casefold(),
                str(item.get("kind", "")).casefold(),
            }:
                return str(item.get("id", ""))
        return self._normalize_graph_folder_id(target)

    def _build_imap_search_criteria(self, config: MailboxConfig) -> list[str]:
        criteria: list[str] = []
        read_state = config.read_state
        if config.unread_only and read_state == DEFAULT_READ_STATE:
            read_state = "unread"
        if read_state == "unread":
            criteria.append("UNSEEN")
        elif read_state == "read":
            criteria.append("SEEN")
        if config.flagged_only:
            criteria.append("FLAGGED")
        return criteria or ["ALL"]

    def _normalize_imap_folder_name(self, folder: str) -> str:
        normalized = str(folder or "").strip()
        if not normalized or normalized.casefold() == "inbox":
            return DEFAULT_FOLDER
        return normalized

    def _resolve_imap_destination_folder(
        self,
        connection: imaplib.IMAP4_SSL,
        destination_folder: str,
        *,
        allow_missing: bool = False,
    ) -> str | None:
        target = str(destination_folder or "").strip()
        if not target:
            if allow_missing:
                return None
            raise MailboxError("缺少目标文件夹", code="invalid_destination_folder")
        folders = self._list_imap_folders_from_connection(connection)
        target_key = target.casefold()
        for item in folders:
            if target_key in {
                str(item.get("id", "")).casefold(),
                str(item.get("name", "")).casefold(),
                str(item.get("display_name", "")).casefold(),
                str(item.get("kind", "")).casefold(),
            }:
                return str(item.get("id", ""))
        if allow_missing:
            return None
        raise MailboxError("目标文件夹不存在", code="invalid_destination_folder")

    def _quote_imap_mailbox(self, folder: str) -> str:
        normalized = self._normalize_imap_folder_name(folder)
        if normalized == DEFAULT_FOLDER:
            return normalized
        escaped = normalized.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{escaped}"'

    def _select_imap_folder(self, connection: imaplib.IMAP4_SSL, folder: str) -> None:
        status, _ = connection.select(self._quote_imap_mailbox(folder))
        if status != "OK":
            raise MailboxError("无法选择邮件文件夹")

    def _logout_imap(self, connection: imaplib.IMAP4_SSL) -> None:
        try:
            connection.logout()
        except Exception:  # pragma: no cover - 容错回收
            pass

    @staticmethod
    def _decode_imap_utf7(value: str) -> str:
        result: list[str] = []
        index = 0
        while index < len(value):
            current = value[index]
            if current != "&":
                result.append(current)
                index += 1
                continue
            end_index = value.find("-", index)
            if end_index == -1:
                result.append(value[index:])
                break
            encoded = value[index + 1 : end_index]
            if not encoded:
                result.append("&")
                index = end_index + 1
                continue
            try:
                chunk = encoded.replace(",", "/")
                padding = "=" * ((4 - len(chunk) % 4) % 4)
                decoded = base64.b64decode(chunk + padding).decode("utf-16-be", "replace")
                result.append(decoded)
            except Exception:
                result.append(f"&{encoded}-")
            index = end_index + 1
        return "".join(result)


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

    def list_folders(self, config: MailboxConfig, method: str) -> list[MailboxFolder]:
        return [self._to_folder(item) for item in self._inner.list_folders(config, method)]

    def list_messages(self, config: MailboxConfig, request: MailboxQuery) -> MessageListResult:
        effective = replace(
            config,
            default_method=request.method,
            top=request.top,
            unread_only=request.unread_only,
            keyword=request.keyword or None,
            folder=request.folder,
            page=request.page,
            page_size=request.page_size,
            read_state=request.read_state,
            has_attachments_only=request.has_attachments_only,
            flagged_only=request.flagged_only,
            importance=request.importance,
            sort_order=request.sort_order,
        )
        raw_messages = self._inner.list_messages(effective, request.method)
        messages = [self._to_summary(item, request.method, request.folder) for item in raw_messages]
        total = len(messages)
        total_pages = (total + request.page_size - 1) // request.page_size if total else 0
        return MessageListResult(
            method=request.method,
            total=total,
            returned=len(messages),
            messages=messages,
            folder=request.folder,
            page=request.page,
            page_size=request.page_size,
            total_pages=total_pages,
            has_prev=request.page > 1 and total_pages > 0,
            has_next=request.page < total_pages,
        )

    def get_message_detail(self, config: MailboxConfig, request: MessageDetailRequest) -> MessageDetail:
        effective = replace(config, default_method=request.method, folder=request.folder)
        raw = self._inner.get_message_detail(effective, request.method, request.message_id)
        return self._to_detail(raw, request.method, request.folder)

    def update_read_state(self, config: MailboxConfig, request: ReadStateUpdateRequest) -> ReadStateUpdateResult:
        effective = replace(config, default_method=request.method, folder=request.folder)
        raw = self._inner.set_read_state(effective, request.method, request.message_id, request.is_read)
        return ReadStateUpdateResult(
            method=raw.get("method", request.method),
            message_id=raw.get("message_id", request.message_id),
            is_read=bool(raw.get("is_read", request.is_read)),
            status=raw.get("status", "updated"),
            source=OutlookMailboxManager.METHOD_LABELS.get(request.method, request.method),
        )

    def update_flag_state(self, config: MailboxConfig, request: FlagStateUpdateRequest) -> FlagStateUpdateResult:
        effective = replace(config, default_method=request.method, folder=request.folder)
        raw = self._inner.set_flag_state(effective, request.method, request.message_id, request.is_flagged)
        return FlagStateUpdateResult(
            method=raw.get("method", request.method),
            message_id=raw.get("message_id", request.message_id),
            is_flagged=bool(raw.get("is_flagged", request.is_flagged)),
            status=raw.get("status", "updated"),
            source=OutlookMailboxManager.METHOD_LABELS.get(request.method, request.method),
        )

    def move_message(self, config: MailboxConfig, request: MessageMoveRequest) -> MessageMoveResult:
        effective = replace(config, default_method=request.method, folder=request.folder)
        raw = self._inner.move_message(effective, request.method, request.message_id, request.destination_folder)
        return MessageMoveResult(
            method=raw.get("method", request.method),
            message_id=raw.get("message_id", request.message_id),
            source_folder=raw.get("source_folder", request.folder),
            destination_folder=raw.get("destination_folder", request.destination_folder),
            status=raw.get("status", "moved"),
            source=OutlookMailboxManager.METHOD_LABELS.get(request.method, request.method),
        )

    def delete_message(self, config: MailboxConfig, request: MessageDeleteRequest) -> MessageDeleteResult:
        effective = replace(config, default_method=request.method, folder=request.folder)
        raw = self._inner.delete_message(effective, request.method, request.message_id)
        return MessageDeleteResult(
            method=raw.get("method", request.method),
            message_id=raw.get("message_id", request.message_id),
            folder=raw.get("folder", request.folder),
            status=raw.get("status", "deleted"),
            source=OutlookMailboxManager.METHOD_LABELS.get(request.method, request.method),
        )

    def save_draft(self, config: MailboxConfig, payload: dict[str, Any]) -> dict[str, Any]:
        method = config.default_method
        effective = replace(config, default_method=method, folder="drafts")
        return self._inner.save_draft(effective, method, payload)

    def send_message(self, config: MailboxConfig, payload: dict[str, Any]) -> dict[str, Any]:
        method = config.default_method
        effective = replace(config, default_method=method)
        return self._inner.send_message(effective, method, payload)

    def reply_message(self, config: MailboxConfig, message_id: str, payload: dict[str, Any], *, reply_all: bool = False) -> dict[str, Any]:
        method = config.default_method
        effective = replace(config, default_method=method)
        return self._inner.reply_message(effective, method, message_id, payload, reply_all=reply_all)

    def forward_message(self, config: MailboxConfig, message_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        method = config.default_method
        effective = replace(config, default_method=method)
        return self._inner.forward_message(effective, method, message_id, payload)

    def upload_attachment(self, config: MailboxConfig, message_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        method = config.default_method
        effective = replace(config, default_method=method)
        return self._inner.upload_attachment(effective, method, message_id, payload)

    def download_attachment(self, config: MailboxConfig, message_id: str, attachment_id: str) -> dict[str, Any]:
        method = config.default_method
        effective = replace(config, default_method=method)
        return self._inner.download_attachment(effective, method, message_id, attachment_id)

    def create_folder(self, config: MailboxConfig, payload: dict[str, Any]) -> dict[str, Any]:
        method = config.default_method
        effective = replace(config, default_method=method)
        return self._inner.create_folder(effective, method, payload)

    def rename_folder(self, config: MailboxConfig, payload: dict[str, Any]) -> dict[str, Any]:
        method = config.default_method
        effective = replace(config, default_method=method)
        return self._inner.rename_folder(effective, method, payload)

    def delete_folder(self, config: MailboxConfig, payload: dict[str, Any]) -> dict[str, Any]:
        method = config.default_method
        effective = replace(config, default_method=method)
        return self._inner.delete_folder(effective, method, payload)

    def resolve_mailbox_email(
        self,
        *,
        client_id: str,
        refresh_token: str,
        proxy: str | None = None,
    ) -> str:
        return self._inner.resolve_mailbox_email(client_id, refresh_token, proxy)

    def _to_folder(self, item: dict[str, Any]) -> MailboxFolder:
        return MailboxFolder(
            id=str(item.get("id", "") or ""),
            name=str(item.get("name", "") or ""),
            display_name=str(item.get("display_name", "") or ""),
            total=int(item.get("total", 0) or 0),
            unread=int(item.get("unread", 0) or 0),
            is_default=bool(item.get("is_default", False)),
            kind=str(item.get("kind", "custom") or "custom"),
        )

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
            folder=item.get("folder", folder),
            internet_message_id=item.get("internet_message_id", ""),
            is_flagged=bool(item.get("is_flagged", False)),
            importance=item.get("importance", "normal"),
            conversation_id=item.get("conversation_id", ""),
        )

    def _to_detail(self, item: dict[str, Any], method: str, folder: str) -> MessageDetail:
        summary = self._to_summary(item, method, folder)
        attachments = [
            AttachmentSummary(
                id=str(attachment.get("id", "") or ""),
                name=str(attachment.get("name", "") or ""),
                content_type=str(attachment.get("content_type", "") or ""),
                size=int(attachment.get("size", 0) or 0),
                is_inline=bool(attachment.get("is_inline", False)),
            )
            for attachment in item.get("attachments", [])
            if isinstance(attachment, dict)
        ]
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
            is_flagged=summary.is_flagged,
            importance=summary.importance,
            conversation_id=item.get("conversation_id", summary.conversation_id),
            body_text=item.get("body_text", item.get("body", "")),
            body_html=item.get("body_html"),
            to_recipients=list(item.get("to", [])),
            cc_recipients=list(item.get("cc", [])),
            bcc_recipients=list(item.get("bcc", [])),
            headers=dict(item.get("headers", {})),
            attachments=attachments,
        )
