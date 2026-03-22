import type {
  Email,
  EmailAccount,
  EmailAttachment,
  Folder,
  MessageMeta as UiMessageMeta,
  MethodValue,
} from '../store/useAppStore';

type JsonRecord = Record<string, unknown>;

interface ApiErrorShape {
  error?: {
    code?: string;
    message?: string;
  };
}

export type MailMethod = MethodValue;

export interface AuthMeResponse {
  authenticated: boolean;
  username: string | null;
}

export interface PaginationMeta {
  page: number;
  page_size: number;
  total: number;
  total_pages: number;
  has_prev: boolean;
  has_next: boolean;
  folder?: string;
}

export interface MailboxSummary {
  id: number;
  label: string;
  email: string;
  preferred_method: MailMethod;
  notes: string;
  created_at?: string;
  updated_at?: string;
}

export interface MailboxProfile extends MailboxSummary {
  client_id: string;
  refresh_token: string;
  proxy?: string | null;
}

export interface FolderItem {
  id: string;
  name: string;
  display_name: string;
  kind: string;
  total: number;
  unread: number;
  is_default?: boolean;
}

export interface AttachmentPayload {
  name: string;
  content_type: string;
  content_base64: string;
}

export interface AttachmentItem {
  id?: string;
  attachment_id?: string;
  name: string;
  content_type: string;
  size: number;
  is_inline: boolean;
  content_base64?: string;
}

export interface MessageMeta extends UiMessageMeta {
  mailbox_id?: number;
  method?: MailMethod;
  message_id?: string;
}

export interface MessageSummary {
  mailbox_id?: number;
  mailbox_label?: string;
  mailbox_email?: string;
  method?: MailMethod;
  message_id: string;
  subject: string;
  sender: string;
  sender_name: string;
  received_at: string;
  is_read: boolean;
  is_flagged: boolean;
  importance: string;
  has_attachments: boolean;
  preview: string;
  folder: string;
  internet_message_id: string;
  conversation_id: string;
  to_recipients?: string[];
  cc_recipients?: string[];
  bcc_recipients?: string[];
  to?: string[];
  cc?: string[];
  bcc?: string[];
  meta?: MessageMeta | null;
}

export interface MessageDetail extends MessageSummary {
  body_text: string;
  body_html?: string | null;
  headers?: Record<string, string>;
  in_reply_to?: string;
  references?: string[];
  attachments: AttachmentItem[];
}

export interface BatchMutationSummary {
  processed: number;
  succeeded: number;
  failed: number;
}

export interface BatchMutationResult extends Record<string, unknown> {
  success?: boolean;
  message?: string;
  mailbox_id?: number;
  message_id?: string;
}

export interface BatchMutationResponse {
  summary: BatchMutationSummary;
  results: BatchMutationResult[];
}

export interface RuleRecord {
  id: number;
  mailbox_id: number;
  name: string;
  enabled: boolean;
  priority: number;
  conditions: Record<string, unknown>;
  actions: Record<string, unknown>;
  created_at: string;
  updated_at: string;
}

export interface RuleApplyResult {
  rule_id: number | null;
  name: string;
  matched: number;
  applied: number;
  failed: number;
}

export interface AuditLogRecord {
  id: number;
  mailbox_id: number | null;
  mailbox_label: string;
  mailbox_email: string;
  actor: string;
  action: string;
  target_type: string;
  target_id: string;
  status: string;
  details: Record<string, unknown>;
  created_at: string;
}

export interface SyncJobRecord {
  id: number;
  mailbox_id: number;
  method: MailMethod;
  requested_by: string;
  scope: Record<string, unknown>;
  status: string;
  processed_messages: number;
  cached_messages: number;
  folders_synced: number;
  error: string;
  started_at: string;
  finished_at: string;
}

export interface SyncStateRecord {
  mailbox_id: number;
  method: MailMethod;
  folder_id: string;
  last_synced_at: string;
  last_message_at: string;
  cached_messages: number;
  status: string;
  error: string;
  updated_at: string;
}

export interface SyncStatusRecord {
  mailbox: MailboxProfile | MailboxSummary;
  jobs: SyncJobRecord[];
  states: SyncStateRecord[];
}

export interface ComposePayload {
  mailboxId: string;
  method: MailMethod;
  draftMessageId?: string;
  messageId?: string;
  subject?: string;
  bodyText?: string;
  bodyHtml?: string;
  toRecipients?: string[];
  ccRecipients?: string[];
  bccRecipients?: string[];
  attachments?: AttachmentPayload[];
  sendNow?: boolean;
}

const JSON_HEADERS = {
  'Content-Type': 'application/json',
};

function normalizeMethod(value: unknown, fallback: MailMethod = 'graph_api'): MailMethod {
  if (value === 'graph_api' || value === 'imap_new' || value === 'imap_old') {
    return value;
  }
  if (value === 'graph') {
    return 'graph_api';
  }
  return fallback;
}

async function requestJson<T>(path: string, init: RequestInit = {}): Promise<T> {
  const headers = new Headers(init.headers ?? {});
  if (init.body && !headers.has('Content-Type')) {
    headers.set('Content-Type', 'application/json');
  }

  const response = await fetch(path, {
    credentials: 'include',
    ...init,
    headers,
  });

  const contentType = response.headers.get('content-type') ?? '';
  const payload = contentType.includes('application/json')
    ? ((await response.json()) as T & ApiErrorShape)
    : null;

  if (!response.ok) {
    const message =
      (payload as ApiErrorShape | null)?.error?.message ??
      `请求失败：${response.status} ${response.statusText}`;
    throw new Error(message);
  }

  return payload as T;
}

function asString(value: unknown, fallback = ''): string {
  return typeof value === 'string' ? value : fallback;
}

function asNumber(value: unknown, fallback = 0): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback;
}

function asBoolean(value: unknown, fallback = false): boolean {
  return typeof value === 'boolean' ? value : fallback;
}

function asStringArray(value: unknown): string[] {
  return Array.isArray(value)
    ? value.filter((item): item is string => typeof item === 'string' && item.trim().length > 0)
    : [];
}

function asRecord(value: unknown): JsonRecord {
  return value && typeof value === 'object' && !Array.isArray(value) ? (value as JsonRecord) : {};
}

function formatHeaders(value: unknown): { headers: string; headersMap: Record<string, string> } {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return { headers: '', headersMap: {} };
  }
  const headersMap = Object.fromEntries(
    Object.entries(value as JsonRecord).map(([key, item]) => [key, String(item ?? '')]),
  );
  return {
    headersMap,
    headers: Object.entries(headersMap)
      .map(([key, item]) => `${key}: ${item}`)
      .join('\n'),
  };
}

function normalizeFolderType(kind: string): Folder['type'] {
  switch (kind.toLowerCase()) {
    case 'inbox':
      return 'inbox';
    case 'sent':
    case 'sentitems':
      return 'sent';
    case 'drafts':
      return 'drafts';
    case 'trash':
    case 'deleteditems':
      return 'trash';
    case 'archive':
      return 'archive';
    case 'junk':
    case 'spam':
      return 'spam';
    default:
      return 'custom';
  }
}

function normalizeImportance(value: unknown): Email['importance'] {
  return value === 'high' || value === 'low' ? value : 'normal';
}

export function methodLabel(method: MailMethod | string): string {
  if (method === 'graph_api' || method === 'graph') {
    return 'Graph API';
  }
  if (method === 'imap_new') {
    return '新版 IMAP';
  }
  if (method === 'imap_old') {
    return '旧版 IMAP';
  }
  return String(method || '未知方式');
}

export function mapMailbox(raw: MailboxSummary | MailboxProfile): EmailAccount {
  const method = normalizeMethod(raw.preferred_method);
  return {
    id: String(raw.id),
    label: raw.label,
    email: raw.email,
    method,
    preferredMethod: method,
    notes: raw.notes ?? '',
    proxy: 'proxy' in raw ? raw.proxy ?? null : null,
    clientId: 'client_id' in raw ? raw.client_id : '',
    refreshToken: 'refresh_token' in raw ? raw.refresh_token : '',
    status: 'unknown',
  };
}

export function mapFolder(raw: FolderItem): Folder {
  const displayName = raw.display_name || raw.name || raw.id;
  return {
    id: raw.id,
    name: raw.name || displayName,
    displayName,
    unreadCount: raw.unread || 0,
    totalCount: raw.total || 0,
    type: normalizeFolderType(raw.kind || displayName),
  };
}

function mapMeta(raw: unknown): UiMessageMeta {
  const source = asRecord(raw);
  const normalizedStatus =
    source.status === 'archived' ||
    source.status === 'deleted' ||
    source.status === 'snoozed' ||
    source.status === 'done'
      ? source.status
      : 'active';
  return {
    tags: asStringArray(source.tags),
    follow_up: asString(source.follow_up),
    notes: asString(source.notes),
    snoozed_until: asString(source.snoozed_until),
    status: normalizedStatus,
    updated_at: asString(source.updated_at),
  };
}

function mapAttachment(raw: AttachmentItem | JsonRecord): EmailAttachment {
  const record = asRecord(raw);
  return {
    id: asString(record.id ?? record.attachment_id),
    name: asString(record.name),
    contentType: asString(record.content_type, 'application/octet-stream'),
    size: asNumber(record.size),
    isInline: asBoolean(record.is_inline),
  };
}

export function mapMessage(raw: MessageSummary | MessageDetail | JsonRecord, methodFallback: MailMethod): Email {
  const record = asRecord(raw);
  const headerPayload = formatHeaders(record.headers);
  return {
    id: asString(record.message_id ?? record.id),
    mailboxId: asString(record.mailbox_id),
    mailboxLabel: asString(record.mailbox_label),
    mailboxEmail: asString(record.mailbox_email),
    subject: asString(record.subject, '无主题'),
    sender: asString(record.sender, '未知发件人'),
    sender_name: asString(record.sender_name),
    to_recipients: asStringArray(record.to_recipients ?? record.to),
    cc_recipients: asStringArray(record.cc_recipients ?? record.cc),
    bcc_recipients: asStringArray(record.bcc_recipients ?? record.bcc),
    date: asString(record.received_at),
    is_read: asBoolean(record.is_read),
    is_flagged: asBoolean(record.is_flagged),
    importance: normalizeImportance(record.importance),
    has_attachments: asBoolean(record.has_attachments),
    preview: asString(record.preview),
    body_text: asString(record.body_text),
    body_html: asString(record.body_html),
    headers: headerPayload.headers,
    headersMap: headerPayload.headersMap,
    internet_message_id: asString(record.internet_message_id),
    conversation_id: asString(record.conversation_id),
    attachments: Array.isArray(record.attachments)
      ? record.attachments.map((item) => mapAttachment(asRecord(item)))
      : [],
    folderId: asString(record.folder, 'INBOX'),
    method: normalizeMethod(record.method, methodFallback),
    meta: mapMeta(record.meta),
    in_reply_to: asString(record.in_reply_to),
    references: asStringArray(record.references),
  };
}

function mapBatchResponse(raw: JsonRecord): BatchMutationResponse {
  return {
    summary: {
      processed: asNumber(asRecord(raw.summary).processed),
      succeeded: asNumber(asRecord(raw.summary).succeeded),
      failed: asNumber(asRecord(raw.summary).failed),
    },
    results: Array.isArray(raw.results) ? raw.results.map((item) => asRecord(item)) : [],
  };
}

function mapPaginationMeta(raw: unknown): PaginationMeta {
  const meta = asRecord(raw);
  return {
    page: asNumber(meta.page, 1),
    page_size: asNumber(meta.page_size, 20),
    total: asNumber(meta.total),
    total_pages: asNumber(meta.total_pages),
    has_prev: asBoolean(meta.has_prev),
    has_next: asBoolean(meta.has_next),
    folder: asString(meta.folder),
  };
}

export async function getAuthMe(): Promise<AuthMeResponse> {
  return requestJson<AuthMeResponse>('/api/auth/me');
}

export async function login(payload: { username: string; password: string }): Promise<AuthMeResponse> {
  return requestJson<AuthMeResponse>('/api/auth/login', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify(payload),
  });
}

export async function logout(): Promise<void> {
  await requestJson('/api/auth/logout', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({}),
  });
}

export async function listMailboxes(params?: { q?: string; page?: number; page_size?: number }): Promise<EmailAccount[]> {
  const query = new URLSearchParams();
  query.set('page', String(params?.page ?? 1));
  query.set('page_size', String(params?.page_size ?? 100));
  if (params?.q) {
    query.set('q', params.q);
  }
  const payload = await requestJson<{ items: MailboxSummary[] }>(`/api/mailboxes?${query.toString()}`);
  return payload.items.map(mapMailbox);
}

export async function createMailbox(payload: {
  label: string;
  email: string;
  client_id: string;
  refresh_token: string;
  preferred_method: MailMethod;
  proxy?: string;
  notes?: string;
}): Promise<EmailAccount> {
  const response = await requestJson<{ mailbox: MailboxProfile }>('/api/mailboxes', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify(payload),
  });
  return mapMailbox(response.mailbox);
}

export async function getMailbox(mailboxId: string): Promise<MailboxProfile> {
  const response = await requestJson<{ mailbox: MailboxProfile }>(`/api/mailboxes/${mailboxId}`);
  return response.mailbox;
}

export async function updateMailbox(mailboxId: string, payload: Record<string, unknown>): Promise<EmailAccount> {
  const response = await requestJson<{ mailbox: MailboxProfile }>(`/api/mailboxes/${mailboxId}`, {
    method: 'PUT',
    headers: JSON_HEADERS,
    body: JSON.stringify(payload),
  });
  return mapMailbox(response.mailbox);
}

export async function batchDeleteMailboxes(mailboxIds: string[]): Promise<BatchMutationResponse> {
  const response = await requestJson<JsonRecord>('/api/mailboxes/delete/batch', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({ mailbox_ids: mailboxIds.map((item) => Number(item)) }),
  });
  return mapBatchResponse(response);
}

export async function batchTestConnections(mailboxIds: string[]): Promise<BatchMutationResponse> {
  const response = await requestJson<JsonRecord>('/api/mailboxes/test-connection/batch', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({ mailbox_ids: mailboxIds.map((item) => Number(item)) }),
  });
  return mapBatchResponse(response);
}

export async function batchUpdatePreferredMethod(mailboxIds: string[], preferredMethod: MailMethod): Promise<BatchMutationResponse> {
  const response = await requestJson<JsonRecord>('/api/mailboxes/preferred-method/batch', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_ids: mailboxIds.map((item) => Number(item)),
      preferred_method: preferredMethod,
    }),
  });
  return mapBatchResponse(response);
}

export async function importMailboxes(rawText: string, preferredMethod: MailMethod): Promise<void> {
  await requestJson('/api/mailboxes/import', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      raw_text: rawText,
      preferred_method: preferredMethod,
    }),
  });
}

export async function listFolders(mailboxId: string, method?: MailMethod): Promise<Folder[]> {
  const payload = await requestJson<{ folders: FolderItem[] }>('/api/mailbox/folders', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(mailboxId),
      ...(method ? { method } : {}),
    }),
  });
  return payload.folders.map(mapFolder);
}

export async function createFolder(params: {
  mailboxId: string;
  method: MailMethod;
  displayName: string;
}): Promise<FolderItem> {
  const response = await requestJson<{ folder: FolderItem }>('/api/mailbox/folder/create', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      display_name: params.displayName,
    }),
  });
  return response.folder;
}

export async function renameFolder(params: {
  mailboxId: string;
  method: MailMethod;
  folderId: string;
  displayName: string;
}): Promise<FolderItem> {
  const response = await requestJson<{ folder: FolderItem }>('/api/mailbox/folder/rename', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      folder_id: params.folderId,
      display_name: params.displayName,
    }),
  });
  return response.folder;
}

export async function deleteFolder(params: {
  mailboxId: string;
  method: MailMethod;
  folderId: string;
}): Promise<FolderItem> {
  const response = await requestJson<{ folder: FolderItem }>('/api/mailbox/folder/delete', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      folder_id: params.folderId,
    }),
  });
  return response.folder;
}

export async function listMessages(params: {
  mailboxId: string;
  method: MailMethod;
  folder: string;
  page: number;
  pageSize: number;
  keyword?: string;
  unreadOnly?: boolean;
  flaggedOnly?: boolean;
  hasAttachmentsOnly?: boolean;
  sortOrder?: 'asc' | 'desc';
}): Promise<{
  items: Email[];
  meta: PaginationMeta;
}> {
  const payload = await requestJson<{ messages: MessageSummary[]; meta: PaginationMeta }>('/api/mailbox/messages', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      folder: params.folder,
      page: params.page,
      page_size: params.pageSize,
      keyword: params.keyword ?? '',
      unread_only: params.unreadOnly ?? false,
      flagged_only: params.flaggedOnly ?? false,
      has_attachments_only: params.hasAttachmentsOnly ?? false,
      sort_order: params.sortOrder ?? 'desc',
    }),
  });
  return {
    items: payload.messages.map((item) => mapMessage(item, params.method)),
    meta: mapPaginationMeta(payload.meta),
  };
}

export async function getMessage(params: {
  mailboxId: string;
  method: MailMethod;
  folder: string;
  messageId: string;
}): Promise<Email> {
  const payload = await requestJson<{ message: MessageDetail }>('/api/mailbox/message', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      folder: params.folder,
      message_id: params.messageId,
    }),
  });
  return mapMessage(payload.message, params.method);
}

export async function updateReadState(params: {
  mailboxId: string;
  method: MailMethod;
  folder: string;
  messageId: string;
  isRead: boolean;
}): Promise<Email> {
  const payload = await requestJson<{ message: MessageDetail }>('/api/mailbox/message/read-state', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      folder: params.folder,
      message_id: params.messageId,
      is_read: params.isRead,
    }),
  });
  return mapMessage(payload.message, params.method);
}

export async function updateFlagState(params: {
  mailboxId: string;
  method: MailMethod;
  folder: string;
  messageId: string;
  isFlagged: boolean;
}): Promise<Email> {
  const payload = await requestJson<{ message: MessageDetail }>('/api/mailbox/message/flag-state', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      folder: params.folder,
      message_id: params.messageId,
      is_flagged: params.isFlagged,
    }),
  });
  return mapMessage(payload.message, params.method);
}

export async function moveMessage(params: {
  mailboxId: string;
  method: MailMethod;
  folder: string;
  messageId: string;
  destinationFolder: string;
}): Promise<void> {
  await requestJson('/api/mailbox/message/move', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      folder: params.folder,
      message_id: params.messageId,
      destination_folder: params.destinationFolder,
    }),
  });
}

export async function deleteMessage(params: {
  mailboxId: string;
  method: MailMethod;
  folder: string;
  messageId: string;
}): Promise<void> {
  await requestJson('/api/mailbox/message/delete', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      folder: params.folder,
      message_id: params.messageId,
    }),
  });
}

export async function batchMessageAction(params: {
  mailboxId: string;
  method: MailMethod;
  folder: string;
  messageIds: string[];
  action: 'mark_read' | 'mark_unread' | 'flag' | 'unflag' | 'delete' | 'archive' | 'move';
  destinationFolder?: string;
}): Promise<BatchMutationResponse> {
  const response = await requestJson<JsonRecord>('/api/mailbox/messages/actions/batch', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      folder: params.folder,
      message_ids: params.messageIds,
      action: params.action,
      ...(params.destinationFolder ? { destination_folder: params.destinationFolder } : {}),
    }),
  });
  return mapBatchResponse(response);
}

function composeRequestBody(payload: ComposePayload): JsonRecord {
  return {
    mailbox_id: Number(payload.mailboxId),
    method: payload.method,
    ...(payload.draftMessageId ? { draft_message_id: payload.draftMessageId } : {}),
    ...(payload.messageId ? { message_id: payload.messageId } : {}),
    subject: payload.subject ?? '',
    body_text: payload.bodyText ?? '',
    ...(payload.bodyHtml ? { body_html: payload.bodyHtml } : {}),
    to_recipients: payload.toRecipients ?? [],
    cc_recipients: payload.ccRecipients ?? [],
    bcc_recipients: payload.bccRecipients ?? [],
    attachments: payload.attachments ?? [],
    send_now: payload.sendNow ?? true,
  };
}

export async function saveDraft(payload: ComposePayload): Promise<Email> {
  const response = await requestJson<{ message: MessageDetail }>('/api/mailbox/message/draft', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify(composeRequestBody(payload)),
  });
  return mapMessage(response.message, payload.method);
}

export async function sendMessage(payload: ComposePayload): Promise<Email> {
  const response = await requestJson<{ message: MessageDetail }>('/api/mailbox/message/send', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify(composeRequestBody(payload)),
  });
  return mapMessage(response.message, payload.method);
}

export async function replyMessage(payload: ComposePayload): Promise<Email> {
  const response = await requestJson<{ message: MessageDetail }>('/api/mailbox/message/reply', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify(composeRequestBody(payload)),
  });
  return mapMessage(response.message, payload.method);
}

export async function replyAllMessage(payload: ComposePayload): Promise<Email> {
  const response = await requestJson<{ message: MessageDetail }>('/api/mailbox/message/reply-all', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify(composeRequestBody(payload)),
  });
  return mapMessage(response.message, payload.method);
}

export async function forwardMessage(payload: ComposePayload): Promise<Email> {
  const response = await requestJson<{ message: MessageDetail }>('/api/mailbox/message/forward', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify(composeRequestBody(payload)),
  });
  return mapMessage(response.message, payload.method);
}

export async function uploadAttachment(params: {
  mailboxId: string;
  method: MailMethod;
  messageId: string;
  attachment: AttachmentPayload;
}): Promise<AttachmentItem> {
  const response = await requestJson<{ attachment: AttachmentItem }>('/api/mailbox/message/attachment/upload', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      message_id: params.messageId,
      ...params.attachment,
    }),
  });
  return response.attachment;
}

export async function downloadAttachment(params: {
  mailboxId: string;
  method: MailMethod;
  folder?: string;
  messageId: string;
  attachmentId: string;
}): Promise<AttachmentItem> {
  const payload = await requestJson<{ attachment: AttachmentItem }>('/api/mailbox/message/attachment/download', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      folder: params.folder ?? '',
      message_id: params.messageId,
      attachment_id: params.attachmentId,
    }),
  });
  return payload.attachment;
}

export async function updateMessageMeta(params: {
  mailboxId: string;
  method: MailMethod;
  folder?: string;
  messageId: string;
  tags?: string[];
  followUp?: string;
  notes?: string;
  snoozedUntil?: string;
  status?: string;
}): Promise<{ meta: UiMessageMeta; message: Email | null }> {
  const payload = await requestJson<{ meta: MessageMeta; message?: MessageDetail | null }>('/api/mailbox/message/meta', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      method: params.method,
      ...(params.folder ? { folder: params.folder } : {}),
      message_id: params.messageId,
      ...(params.tags ? { tags: params.tags } : {}),
      ...(params.followUp !== undefined ? { follow_up: params.followUp } : {}),
      ...(params.notes !== undefined ? { notes: params.notes } : {}),
      ...(params.snoozedUntil !== undefined ? { snoozed_until: params.snoozedUntil } : {}),
      ...(params.status !== undefined ? { status: params.status } : {}),
    }),
  });
  return {
    meta: mapMeta(payload.meta),
    message: payload.message ? mapMessage(payload.message, params.method) : null,
  };
}

export async function searchMessages(params: {
  query?: string;
  mailboxIds?: string[];
  method?: MailMethod;
  folder?: string;
  tag?: string;
  tags?: string[];
  unreadOnly?: boolean;
  flaggedOnly?: boolean;
  hasAttachmentsOnly?: boolean;
  includeSnoozed?: boolean;
  page?: number;
  pageSize?: number;
  sortOrder?: 'asc' | 'desc';
}): Promise<{ items: Email[]; meta: PaginationMeta }> {
  const payload = await requestJson<{ items: MessageSummary[]; meta: PaginationMeta }>('/api/mailboxes/messages/search', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      query: params.query ?? '',
      ...(params.mailboxIds?.length ? { mailbox_ids: params.mailboxIds.map((item) => Number(item)) } : {}),
      ...(params.method ? { method: params.method } : {}),
      ...(params.folder ? { folder: params.folder } : {}),
      ...(params.tag ? { tag: params.tag } : {}),
      ...(params.tags?.length ? { tags: params.tags } : {}),
      unread_only: params.unreadOnly ?? false,
      flagged_only: params.flaggedOnly ?? false,
      has_attachments_only: params.hasAttachmentsOnly ?? false,
      include_snoozed: params.includeSnoozed ?? true,
      page: params.page ?? 1,
      page_size: params.pageSize ?? 20,
      sort_order: params.sortOrder ?? 'desc',
    }),
  });
  return {
    items: payload.items.map((item) => mapMessage(item, normalizeMethod(item.method))),
    meta: mapPaginationMeta(payload.meta),
  };
}

export async function getThread(params: {
  mailboxId: string;
  method?: MailMethod;
  folder?: string;
  messageId?: string;
  conversationId?: string;
}): Promise<Email[]> {
  const payload = await requestJson<{ items: MessageDetail[] }>('/api/mailbox/thread', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      ...(params.method ? { method: params.method } : {}),
      ...(params.folder ? { folder: params.folder } : {}),
      ...(params.messageId ? { message_id: params.messageId } : {}),
      ...(params.conversationId ? { conversation_id: params.conversationId } : {}),
    }),
  });
  return payload.items.map((item) => mapMessage(item, params.method ?? normalizeMethod(item.method)));
}

export async function listRules(params: {
  mailboxId: string;
  enabledOnly?: boolean;
}): Promise<RuleRecord[]> {
  const payload = await requestJson<{ items: RuleRecord[] }>('/api/mailbox/rules/list', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      enabled_only: params.enabledOnly ?? false,
    }),
  });
  return payload.items;
}

export async function createRule(params: {
  mailboxId: string;
  name: string;
  enabled: boolean;
  priority: number;
  conditions: Record<string, unknown>;
  actions: Record<string, unknown>;
}): Promise<RuleRecord> {
  const payload = await requestJson<{ rule: RuleRecord }>('/api/mailbox/rules', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      name: params.name,
      enabled: params.enabled,
      priority: params.priority,
      conditions: params.conditions,
      actions: params.actions,
    }),
  });
  return payload.rule;
}

export async function updateRule(params: {
  mailboxId: string;
  ruleId: number;
  name?: string;
  enabled?: boolean;
  priority?: number;
  conditions?: Record<string, unknown>;
  actions?: Record<string, unknown>;
}): Promise<RuleRecord> {
  const payload = await requestJson<{ rule: RuleRecord }>('/api/mailbox/rules/update', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      rule_id: params.ruleId,
      ...(params.name !== undefined ? { name: params.name } : {}),
      ...(params.enabled !== undefined ? { enabled: params.enabled } : {}),
      ...(params.priority !== undefined ? { priority: params.priority } : {}),
      ...(params.conditions !== undefined ? { conditions: params.conditions } : {}),
      ...(params.actions !== undefined ? { actions: params.actions } : {}),
    }),
  });
  return payload.rule;
}

export async function deleteRule(params: { mailboxId: string; ruleId: number }): Promise<void> {
  await requestJson('/api/mailbox/rules/delete', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      rule_id: params.ruleId,
    }),
  });
}

export async function applyRules(params: {
  mailboxId: string;
  method?: MailMethod;
  folder?: string;
  ruleId?: number;
  limit?: number;
}): Promise<{ results: RuleApplyResult[]; count: number }> {
  return requestJson<{ results: RuleApplyResult[]; count: number }>('/api/mailbox/rules/apply', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      mailbox_id: Number(params.mailboxId),
      ...(params.method ? { method: params.method } : {}),
      ...(params.folder ? { folder: params.folder } : {}),
      ...(params.ruleId ? { rule_id: params.ruleId } : {}),
      ...(params.limit ? { limit: params.limit } : {}),
    }),
  });
}

export async function listAuditLogs(params?: {
  mailboxId?: string;
  action?: string;
  page?: number;
  pageSize?: number;
}): Promise<{ items: AuditLogRecord[]; meta: PaginationMeta }> {
  const payload = await requestJson<{ items: AuditLogRecord[]; meta: PaginationMeta }>('/api/audit/logs', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      ...(params?.mailboxId ? { mailbox_id: Number(params.mailboxId) } : {}),
      ...(params?.action ? { action: params.action } : {}),
      page: params?.page ?? 1,
      page_size: params?.pageSize ?? 20,
    }),
  });
  return {
    items: payload.items,
    meta: mapPaginationMeta(payload.meta),
  };
}

export async function runSync(params: {
  mailboxId?: string;
  mailboxIds?: string[];
  method?: MailMethod;
  folderLimit?: number;
  messageLimit?: number;
  includeBody?: boolean;
  applyRules?: boolean;
}): Promise<{ results: Array<{ mailbox: MailboxSummary; job: SyncJobRecord }>; count: number }> {
  return requestJson('/api/mailbox/sync/run', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      ...(params.mailboxId ? { mailbox_id: Number(params.mailboxId) } : {}),
      ...(params.mailboxIds?.length ? { mailbox_ids: params.mailboxIds.map((item) => Number(item)) } : {}),
      ...(params.method ? { method: params.method } : {}),
      folder_limit: params.folderLimit ?? 5,
      message_limit: params.messageLimit ?? 20,
      include_body: params.includeBody ?? true,
      apply_rules: params.applyRules ?? true,
    }),
  });
}

export async function getSyncStatus(params?: {
  mailboxId?: string;
  mailboxIds?: string[];
  method?: MailMethod;
}): Promise<{ items: SyncStatusRecord[]; count: number }> {
  return requestJson('/api/mailbox/sync/status', {
    method: 'POST',
    headers: JSON_HEADERS,
    body: JSON.stringify({
      ...(params?.mailboxId ? { mailbox_id: Number(params.mailboxId) } : {}),
      ...(params?.mailboxIds?.length ? { mailbox_ids: params.mailboxIds.map((item) => Number(item)) } : {}),
      ...(params?.method ? { method: params.method } : {}),
    }),
  });
}

export const api = {
  getAuthMe,
  login,
  logout,
  listMailboxes,
  createMailbox,
  getMailbox,
  updateMailbox,
  batchDeleteMailboxes,
  batchTestConnections,
  batchUpdatePreferredMethod,
  importMailboxes,
  listFolders,
  createFolder,
  renameFolder,
  deleteFolder,
  listMessages,
  getMessage,
  updateReadState,
  updateFlagState,
  moveMessage,
  deleteMessage,
  batchMessageAction,
  saveDraft,
  sendMessage,
  replyMessage,
  replyAllMessage,
  forwardMessage,
  uploadAttachment,
  downloadAttachment,
  updateMessageMeta,
  searchMessages,
  getThread,
  listRules,
  createRule,
  updateRule,
  deleteRule,
  applyRules,
  listAuditLogs,
  runSync,
  getSyncStatus,
};
