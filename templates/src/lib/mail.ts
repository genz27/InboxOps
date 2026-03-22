import type {
  Email,
  EmailAccount,
  EmailAttachment,
  Folder,
  MessageMeta,
  MethodValue,
} from '../store/useAppStore';

type MailMethod = MethodValue;

interface MailboxSummary {
  id: number | string;
  label?: string;
  email: string;
  preferred_method?: MailMethod;
  method?: MailMethod;
  notes?: string;
  remark?: string;
}

interface MailboxProfile extends MailboxSummary {
  proxy?: string | null;
  client_id?: string;
  refresh_token?: string;
}

interface FolderItem {
  id: string;
  name?: string;
  display_name?: string;
  unread?: number;
  total?: number;
  kind?: string;
  path?: string;
  parent_id?: string | null;
  is_system?: boolean;
}

interface AttachmentItem {
  id?: string;
  attachment_id?: string;
  name: string;
  content_type?: string;
  type?: string;
  size?: number;
  is_inline?: boolean;
  isInline?: boolean;
  content_id?: string;
}

interface MessageBase {
  message_id?: string;
  id?: string;
  mailbox_label?: string;
  mailbox_email?: string;
  subject?: string;
  sender?: string;
  sender_name?: string;
  to_recipients?: string[];
  to?: string[];
  cc_recipients?: string[];
  cc?: string[];
  bcc_recipients?: string[];
  bcc?: string[];
  received_at?: string;
  is_read?: boolean;
  is_flagged?: boolean;
  importance?: Email['importance'];
  has_attachments?: boolean;
  preview?: string;
  body?: string;
  body_text?: string;
  body_html?: string;
  headers?: Record<string, string> | string;
  internet_message_id?: string;
  conversation_id?: string;
  attachments?: AttachmentItem[];
  folder?: string;
  method?: MailMethod;
  meta?: MessageMeta | null;
}

interface MessageSummary extends MessageBase {}

interface MessageDetail extends MessageBase {}

function normalizeMethod(method: MailMethod | undefined): MethodValue {
  if (method === 'imap_new') {
    return 'imap_new';
  }
  if (method === 'imap_old' || method === 'IMAP') {
    return 'imap_old';
  }
  return 'graph_api';
}

export function methodLabel(method: MailMethod): string {
  const normalized = normalizeMethod(method);
  return {
    graph_api: 'Graph API',
    imap_new: 'IMAP New',
    imap_old: 'IMAP Old',
  }[normalized];
}

export function emptyMeta(): MessageMeta {
  return {
    tags: [],
    follow_up: '',
    notes: '',
    snoozed_until: '',
    status: 'active',
  };
}

export function mapMailboxToAccount(mailbox: MailboxSummary | MailboxProfile): EmailAccount {
  const method = normalizeMethod(mailbox.preferred_method || mailbox.method);
  const notes = mailbox.notes || mailbox.remark || '';

  return {
    id: String(mailbox.id),
    label: mailbox.label || mailbox.email,
    email: mailbox.email,
    method,
    preferredMethod: method,
    status: 'unknown',
    notes,
    remark: notes,
    proxy: 'proxy' in mailbox ? mailbox.proxy ?? null : null,
    clientId: 'client_id' in mailbox ? mailbox.client_id || '' : '',
    refreshToken: 'refresh_token' in mailbox ? mailbox.refresh_token || '' : '',
  };
}

export function mapFolderToUi(folder: FolderItem): Folder {
  const kind = (folder.kind || 'custom').toLowerCase();
  const normalizedType: Folder['type'] =
    kind === 'inbox'
      ? 'inbox'
      : kind === 'sent'
        ? 'sent'
        : kind === 'drafts'
          ? 'drafts'
          : kind === 'trash'
            ? 'trash'
            : kind === 'junk'
              ? 'spam'
              : kind === 'archive'
                ? 'archive'
                : 'custom';

  const displayName = folder.display_name || folder.name || folder.id;
  return {
    id: folder.id,
    name: folder.name || displayName,
    displayName,
    unreadCount: folder.unread || 0,
    totalCount: folder.total || 0,
    type: normalizedType,
    path: folder.path,
    parentId: folder.parent_id ?? null,
    isSystem: folder.is_system,
  };
}

function mapAttachment(attachment: AttachmentItem): EmailAttachment {
  const contentType = attachment.content_type || attachment.type || 'application/octet-stream';
  const isInline = Boolean(attachment.is_inline ?? attachment.isInline);

  return {
    id: attachment.id || attachment.attachment_id || attachment.name,
    name: attachment.name,
    contentType,
    type: contentType,
    size: attachment.size || 0,
    isInline,
    is_inline: isInline,
    contentId: attachment.content_id,
  };
}

function normalizeHeaders(headers: MessageBase['headers']): { headers: string; headersMap?: Record<string, string> } {
  if (typeof headers === 'string') {
    return { headers };
  }
  if (!headers || typeof headers !== 'object' || Array.isArray(headers)) {
    return { headers: '', headersMap: {} };
  }

  const headersMap = headers as Record<string, string>;
  return {
    headers: Object.entries(headersMap)
      .map(([key, value]) => `${key}: ${value}`)
      .join('\n'),
    headersMap,
  };
}

export function mapMessageToEmail(
  message: MessageSummary | MessageDetail,
  mailboxId: number,
  fallbackMeta?: MessageMeta | null
): Email {
  const normalizedMethod = normalizeMethod(message.method);
  const normalizedHeaders = normalizeHeaders(message.headers);
  const bodyText = message.body_text || message.body || '';
  const bodyHtml = message.body_html || '';

  return {
    id: message.message_id || message.id || '',
    mailboxId: String(mailboxId),
    mailboxLabel: message.mailbox_label,
    mailboxEmail: message.mailbox_email,
    subject: message.subject || '无主题',
    sender: message.sender || '未知发件人',
    sender_name: message.sender_name || '',
    to_recipients: Array.isArray(message.to_recipients) ? message.to_recipients : Array.isArray(message.to) ? message.to : [],
    cc_recipients: Array.isArray(message.cc_recipients) ? message.cc_recipients : Array.isArray(message.cc) ? message.cc : [],
    bcc_recipients: Array.isArray(message.bcc_recipients) ? message.bcc_recipients : Array.isArray(message.bcc) ? message.bcc : [],
    date: message.received_at || '',
    is_read: Boolean(message.is_read),
    is_flagged: Boolean(message.is_flagged),
    importance: message.importance || 'normal',
    has_attachments: Boolean(message.has_attachments),
    preview: message.preview || '',
    body_text: bodyText,
    body_html: bodyHtml,
    headers: normalizedHeaders.headers,
    headersMap: normalizedHeaders.headersMap,
    internet_message_id: message.internet_message_id || '',
    conversation_id: message.conversation_id || '',
    attachments: Array.isArray(message.attachments) ? message.attachments.map(mapAttachment) : [],
    folderId: message.folder || 'INBOX',
    method: normalizedMethod,
    meta: message.meta || fallbackMeta || emptyMeta(),
  };
}

export async function fileToAttachmentPayload(file: File) {
  const contentBase64 = await readFileAsBase64(file);
  return {
    name: file.name,
    content_type: file.type || 'application/octet-stream',
    content_base64: contentBase64,
  };
}

export function toDatetimeLocalValue(value: string | null | undefined): string {
  if (!value) {
    return '';
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '';
  }
  const offset = date.getTimezoneOffset();
  const adjusted = new Date(date.getTime() - offset * 60000);
  return adjusted.toISOString().slice(0, 16);
}

export function fromDatetimeLocalValue(value: string): string {
  if (!value) {
    return '';
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? '' : date.toISOString();
}

function readFileAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = typeof reader.result === 'string' ? reader.result : '';
      const base64 = result.includes(',') ? result.split(',', 2)[1] : result;
      resolve(base64);
    };
    reader.onerror = () => reject(reader.error || new Error('文件读取失败'));
    reader.readAsDataURL(file);
  });
}
