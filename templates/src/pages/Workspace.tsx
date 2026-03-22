import React, { useDeferredValue, useEffect, useMemo, useState } from 'react';
import {
  AlertCircle,
  Archive,
  ChevronLeft,
  ChevronRight,
  Clock,
  Download,
  FilePlus2,
  FolderPen,
  FolderPlus,
  Forward,
  LoaderCircle,
  Mail,
  MailOpen,
  Paperclip,
  RefreshCw,
  Reply,
  ReplyAll,
  Search,
  ShieldCheck,
  Star,
  Trash2,
  X,
} from 'lucide-react';
import { format } from 'date-fns';
import { useI18n } from '../i18n';
import { Button } from '../components/ui/Button';
import { Input } from '../components/ui/Input';
import { Badge } from '../components/ui/Badge';
import { cn } from '../lib/utils';
import {
  type AuditLogRecord,
  type AttachmentPayload,
  type PaginationMeta,
  type RuleRecord,
  type SyncStatusRecord,
  applyRules,
  batchMessageAction,
  createFolder,
  createRule,
  deleteFolder,
  deleteMessage,
  deleteRule,
  downloadAttachment,
  forwardMessage,
  getMessage,
  getThread,
  getSyncStatus,
  listAuditLogs,
  listFolders,
  listMailboxes,
  listMessages,
  listRules,
  renameFolder,
  replyAllMessage,
  replyMessage,
  runSync,
  saveDraft,
  searchMessages,
  sendMessage,
  updateFlagState,
  updateMessageMeta,
  updateReadState,
  updateRule,
} from '../lib/api';
import { emptyMeta, fileToAttachmentPayload, fromDatetimeLocalValue, methodLabel, toDatetimeLocalValue } from '../lib/mail';
import { useAppStore, type Email, type EmailAccount, type Folder, type MessageMeta, type MethodValue } from '../store/useAppStore';

const ITEMS_PER_PAGE = 20;
const MAIL_LIST_AUTO_REFRESH_MS = 30000;
const EMPTY_PAGINATION: PaginationMeta = {
  page: 1,
  page_size: ITEMS_PER_PAGE,
  total: 0,
  total_pages: 0,
  has_prev: false,
  has_next: false,
};

type ViewMode = 'html' | 'text' | 'headers';
type ComposeMode = 'new' | 'reply' | 'replyAll' | 'forward';

interface ComposeFormState {
  open: boolean;
  mode: ComposeMode;
  messageId: string;
  draftMessageId: string;
  subject: string;
  to: string;
  cc: string;
  bcc: string;
  bodyText: string;
  attachments: AttachmentPayload[];
  attachmentNames: string[];
  submitting: boolean;
}

interface MetaFormState {
  tags: string;
  followUp: string;
  notes: string;
  snoozedUntil: string;
  status: string;
}

interface RuleEditorState {
  id: number | null;
  name: string;
  enabled: boolean;
  priority: number;
  conditionsText: string;
  actionsText: string;
}

interface PendingFocus {
  mailboxId: string;
  folderId: string;
  messageId: string;
}

const EMPTY_COMPOSE: ComposeFormState = {
  open: false,
  mode: 'new',
  messageId: '',
  draftMessageId: '',
  subject: '',
  to: '',
  cc: '',
  bcc: '',
  bodyText: '',
  attachments: [],
  attachmentNames: [],
  submitting: false,
};

const EMPTY_META_FORM: MetaFormState = {
  tags: '',
  followUp: '',
  notes: '',
  snoozedUntil: '',
  status: 'active',
};

const EMPTY_RULE_EDITOR: RuleEditorState = {
  id: null,
  name: '',
  enabled: true,
  priority: 100,
  conditionsText: '{\n  "subject_contains": ""\n}',
  actionsText: '{\n  "mark_read": true,\n  "tags": [""]\n}',
};

function splitRecipients(value: string): string[] {
  return value
    .split(/[\n,;]+/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function joinRecipients(items: string[]): string {
  return items.join(', ');
}

function makeQuotedBody(message: Email): string {
  const receivedAt = message.date ? format(new Date(message.date), 'yyyy-MM-dd HH:mm') : '';
  const quoted = message.body_text || message.preview || '';
  return `\n\n\n--- 原始邮件 ---\n主题: ${message.subject}\n发件人: ${message.sender}\n时间: ${receivedAt}\n\n${quoted}`;
}

function titleWithPrefix(prefix: string, subject: string): string {
  const trimmedSubject = subject.trim() || '无主题';
  return trimmedSubject.toLowerCase().startsWith(prefix.toLowerCase()) ? trimmedSubject : `${prefix}${trimmedSubject}`;
}

function base64ToBlob(base64: string, contentType: string): Blob {
  const binary = window.atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new Blob([bytes], { type: contentType || 'application/octet-stream' });
}

function triggerFileDownload(filename: string, blob: Blob) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function parseJsonObject(text: string, fieldName: string): Record<string, unknown> {
  try {
    const value = JSON.parse(text);
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
      throw new Error();
    }
    return value as Record<string, unknown>;
  } catch {
    throw new Error(`${fieldName} 必须是合法 JSON 对象`);
  }
}

function isDraftFolder(email: Email | null, folderId: string) {
  if (!email) {
    return false;
  }
  const normalized = (email.folderId || folderId || '').toLowerCase();
  return normalized.includes('draft');
}

export default function Workspace() {
  const { t } = useI18n();
  const syncAccounts = useAppStore((state) => state.setAccounts);
  const syncActiveMailboxId = useAppStore((state) => state.setActiveMailboxId);
  const [accounts, setAccounts] = useState<EmailAccount[]>([]);
  const [activeMailboxId, setActiveMailboxId] = useState('');
  const [folders, setFolders] = useState<Folder[]>([]);
  const [messages, setMessages] = useState<Email[]>([]);
  const [listMeta, setListMeta] = useState<PaginationMeta>(EMPTY_PAGINATION);
  const [selectedMessageIds, setSelectedMessageIds] = useState<string[]>([]);
  const [activeFolderId, setActiveFolderId] = useState('');
  const [activeMessageId, setActiveMessageId] = useState('');
  const [activeMessage, setActiveMessage] = useState<Email | null>(null);
  const [threadItems, setThreadItems] = useState<Email[]>([]);
  const [auditLogs, setAuditLogs] = useState<AuditLogRecord[]>([]);
  const [syncItems, setSyncItems] = useState<SyncStatusRecord[]>([]);
  const [rules, setRules] = useState<RuleRecord[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const deferredSearchQuery = useDeferredValue(searchQuery);
  const [filterUnread, setFilterUnread] = useState(false);
  const [filterStarred, setFilterStarred] = useState(false);
  const [filterAttachment, setFilterAttachment] = useState(false);
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('desc');
  const [page, setPage] = useState(1);
  const [viewMode, setViewMode] = useState<ViewMode>('html');
  const [loadingAccounts, setLoadingAccounts] = useState(false);
  const [loadingFolders, setLoadingFolders] = useState(false);
  const [loadingMessages, setLoadingMessages] = useState(false);
  const [loadingDetail, setLoadingDetail] = useState(false);
  const [loadingOps, setLoadingOps] = useState(false);
  const [loadingRules, setLoadingRules] = useState(false);
  const [busyAction, setBusyAction] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [foldersCollapsed, setFoldersCollapsed] = useState(true);
  const [detailSidebarCollapsed, setDetailSidebarCollapsed] = useState(true);
  const [metaPanelCollapsed, setMetaPanelCollapsed] = useState(false);
  const [threadPanelCollapsed, setThreadPanelCollapsed] = useState(false);
  const [compose, setCompose] = useState<ComposeFormState>(EMPTY_COMPOSE);
  const [metaForm, setMetaForm] = useState<MetaFormState>(EMPTY_META_FORM);
  const [searchOpen, setSearchOpen] = useState(false);
  const [searchAcrossAll, setSearchAcrossAll] = useState(true);
  const [globalSearchQuery, setGlobalSearchQuery] = useState('');
  const [globalSearchResults, setGlobalSearchResults] = useState<Email[]>([]);
  const [globalSearchMeta, setGlobalSearchMeta] = useState<PaginationMeta>(EMPTY_PAGINATION);
  const [globalSearchLoading, setGlobalSearchLoading] = useState(false);
  const [rulesOpen, setRulesOpen] = useState(false);
  const [opsOpen, setOpsOpen] = useState(false);
  const [ruleEditor, setRuleEditor] = useState<RuleEditorState>(EMPTY_RULE_EDITOR);
  const [pendingFocus, setPendingFocus] = useState<PendingFocus | null>(null);

  const activeMailbox = useMemo(
    () => accounts.find((account) => account.id === activeMailboxId) ?? null,
    [accounts, activeMailboxId],
  );
  const activeMethod: MethodValue = activeMailbox?.preferredMethod ?? 'graph_api';
  const activeFolder = useMemo(() => folders.find((folder) => folder.id === activeFolderId) ?? null, [folders, activeFolderId]);
  const allVisibleSelected = messages.length > 0 && messages.every((message) => selectedMessageIds.includes(message.id));
  const syncStatus = syncItems[0] ?? null;

  useEffect(() => {
    syncAccounts(accounts);
  }, [accounts, syncAccounts]);

  useEffect(() => {
    syncActiveMailboxId(activeMailboxId || null);
  }, [activeMailboxId, syncActiveMailboxId]);

  async function loadAccounts() {
    setLoadingAccounts(true);
    setError('');
    try {
      const items = await listMailboxes();
      setAccounts(items);
      setActiveMailboxId((current) => {
        if (items.some((item) => item.id === current)) {
          return current;
        }
        return items[0]?.id ?? '';
      });
      if (items.length === 0) {
        setFolders([]);
        setMessages([]);
        setActiveFolderId('');
        setActiveMessageId('');
        setActiveMessage(null);
      }
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : '加载邮箱档案失败');
    } finally {
      setLoadingAccounts(false);
    }
  }

  async function loadFoldersForMailbox(mailboxId: string, method: MethodValue) {
    setLoadingFolders(true);
    try {
      const items = await listFolders(mailboxId, method);
      setFolders(items);
      setActiveFolderId((current) => {
        if (items.some((item) => item.id === current)) {
          return current;
        }
        return items.find((item) => item.type === 'inbox')?.id ?? items[0]?.id ?? 'INBOX';
      });
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : '加载文件夹失败');
      setFolders([]);
      setActiveFolderId('');
    } finally {
      setLoadingFolders(false);
    }
  }

  async function loadMessagesForFolder(nextPage = page, options?: { silent?: boolean }) {
    if (!activeMailbox || !activeFolderId) {
      return;
    }
    const silent = options?.silent === true;
    if (!silent) {
      setLoadingMessages(true);
      setError('');
    }
    try {
      const response = await listMessages({
        mailboxId: activeMailbox.id,
        method: activeMethod,
        folder: activeFolderId,
        page: nextPage,
        pageSize: ITEMS_PER_PAGE,
        keyword: deferredSearchQuery,
        unreadOnly: filterUnread,
        flaggedOnly: filterStarred,
        hasAttachmentsOnly: filterAttachment,
        sortOrder,
      });
      setMessages(response.items);
      setListMeta(response.meta);
      setSelectedMessageIds((current) => current.filter((item) => response.items.some((message) => message.id === item)));
      if (!response.items.some((message) => message.id === activeMessageId) && !pendingFocus) {
        setActiveMessageId('');
        setActiveMessage(null);
        setThreadItems([]);
      }
    } catch (requestError) {
      if (!silent) {
        setError(requestError instanceof Error ? requestError.message : '加载邮件失败');
        setMessages([]);
        setListMeta(EMPTY_PAGINATION);
      }
    } finally {
      if (!silent) {
        setLoadingMessages(false);
      }
    }
  }

  async function loadMessageDetail(messageId: string, folderId = activeFolderId) {
    if (!activeMailbox || !messageId) {
      return;
    }
    setLoadingDetail(true);
    setError('');
    try {
      const message = await getMessage({
        mailboxId: activeMailbox.id,
        method: activeMethod,
        folder: folderId,
        messageId,
      });
      setActiveMessageId(messageId);
      setActiveMessage(message);
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : '加载邮件详情失败');
    } finally {
      setLoadingDetail(false);
    }
  }

  async function loadThreadForMessage(message: Email | null) {
    if (!activeMailbox || !message) {
      setThreadItems([]);
      return;
    }
    try {
      const items = await getThread({
        mailboxId: activeMailbox.id,
        method: activeMethod,
        folder: message.folderId,
        messageId: message.id,
        conversationId: message.conversation_id || undefined,
      });
      setThreadItems(items);
    } catch {
      setThreadItems([]);
    }
  }

  async function loadOpsForMailbox(mailboxId: string) {
    setLoadingOps(true);
    try {
      const [auditPayload, syncPayload] = await Promise.all([
        listAuditLogs({ mailboxId, page: 1, pageSize: 12 }),
        getSyncStatus({ mailboxId }),
      ]);
      setAuditLogs(auditPayload.items);
      setSyncItems(syncPayload.items);
    } catch {
      setAuditLogs([]);
      setSyncItems([]);
    } finally {
      setLoadingOps(false);
    }
  }

  async function loadRulesForMailbox(mailboxId: string) {
    setLoadingRules(true);
    try {
      setRules(await listRules({ mailboxId }));
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : '加载规则失败');
      setRules([]);
    } finally {
      setLoadingRules(false);
    }
  }

  useEffect(() => {
    void loadAccounts();
  }, []);

  useEffect(() => {
    if (!activeMailbox) {
      return;
    }
    void loadFoldersForMailbox(activeMailbox.id, activeMethod);
    void loadOpsForMailbox(activeMailbox.id);
  }, [activeMailbox?.id, activeMethod]);

  useEffect(() => {
    if (!activeMailbox || !activeFolderId) {
      return;
    }
    void loadMessagesForFolder(page);
  }, [activeMailbox?.id, activeMethod, activeFolderId, page, deferredSearchQuery, filterUnread, filterStarred, filterAttachment, sortOrder]);

  useEffect(() => {
    if (!activeMailbox || !activeFolderId) {
      return;
    }

    const timerId = window.setInterval(() => {
      if (document.visibilityState === 'hidden' || loadingMessages || loadingDetail || busyAction !== '') {
        return;
      }
      void loadMessagesForFolder(page, { silent: true });
    }, MAIL_LIST_AUTO_REFRESH_MS);

    return () => {
      window.clearInterval(timerId);
    };
  }, [
    activeMailbox?.id,
    activeMethod,
    activeFolderId,
    page,
    deferredSearchQuery,
    filterUnread,
    filterStarred,
    filterAttachment,
    sortOrder,
    loadingMessages,
    loadingDetail,
    busyAction,
  ]);

  useEffect(() => {
    void loadThreadForMessage(activeMessage);
    if (!activeMessage) {
      setMetaForm(EMPTY_META_FORM);
      return;
    }
    const meta = activeMessage.meta ?? emptyMeta();
    setMetaForm({
      tags: meta.tags.join(', '),
      followUp: meta.follow_up,
      notes: meta.notes,
      snoozedUntil: toDatetimeLocalValue(meta.snoozed_until),
      status: meta.status || 'active',
    });
  }, [activeMessage]);

  useEffect(() => {
    if (!rulesOpen || !activeMailbox) {
      return;
    }
    void loadRulesForMailbox(activeMailbox.id);
  }, [rulesOpen, activeMailbox?.id]);

  useEffect(() => {
    if (!pendingFocus || !activeMailbox) {
      return;
    }
    if (pendingFocus.mailboxId !== activeMailbox.id) {
      return;
    }
    if (pendingFocus.folderId !== activeFolderId) {
      setActiveFolderId(pendingFocus.folderId);
      setPage(1);
      return;
    }
    void (async () => {
      await loadMessageDetail(pendingFocus.messageId, pendingFocus.folderId);
      setPendingFocus(null);
    })();
  }, [pendingFocus, activeMailbox?.id, activeFolderId]);

  async function refreshEverything() {
    if (!activeMailbox) {
      await loadAccounts();
      return;
    }
    await Promise.all([
      loadFoldersForMailbox(activeMailbox.id, activeMethod),
      loadMessagesForFolder(1),
      loadOpsForMailbox(activeMailbox.id),
    ]);
    setPage(1);
  }

  function updateNotice(message: string) {
    setNotice(message);
    setError('');
  }

  function updateError(requestError: unknown, fallback: string) {
    setError(requestError instanceof Error ? requestError.message : fallback);
    setNotice('');
  }

  function updateMessageCollection(nextMessage: Email) {
    setMessages((current) => current.map((item) => (item.id === nextMessage.id ? { ...item, ...nextMessage } : item)));
    setActiveMessage(nextMessage);
  }

  async function handleSingleStateUpdate(type: 'read' | 'flag', value: boolean) {
    if (!activeMailbox || !activeMessage) {
      return;
    }
    setBusyAction(type);
    try {
      const nextMessage =
        type === 'read'
          ? await updateReadState({
              mailboxId: activeMailbox.id,
              method: activeMethod,
              folder: activeMessage.folderId,
              messageId: activeMessage.id,
              isRead: value,
            })
          : await updateFlagState({
              mailboxId: activeMailbox.id,
              method: activeMethod,
              folder: activeMessage.folderId,
              messageId: activeMessage.id,
              isFlagged: value,
            });
      updateMessageCollection(nextMessage);
      updateNotice('邮件状态已更新');
    } catch (requestError) {
      updateError(requestError, '更新邮件状态失败');
    } finally {
      setBusyAction('');
    }
  }

  async function handleBatchAction(
    action: 'mark_read' | 'mark_unread' | 'flag' | 'unflag' | 'delete' | 'archive' | 'move',
    messageIds: string[],
    destinationFolder?: string,
  ) {
    if (!activeMailbox || !activeFolderId || messageIds.length === 0) {
      return;
    }
    setBusyAction(action);
    try {
      const payload = await batchMessageAction({
        mailboxId: activeMailbox.id,
        method: activeMethod,
        folder: activeFolderId,
        messageIds,
        action,
        destinationFolder,
      });
      if (messageIds.includes(activeMessageId)) {
        setActiveMessage(null);
        setActiveMessageId('');
      }
      setSelectedMessageIds((current) => current.filter((item) => !messageIds.includes(item)));
      updateNotice(`操作完成，成功 ${payload.summary.succeeded} 条`);
      await loadMessagesForFolder(page);
      await loadOpsForMailbox(activeMailbox.id);
    } catch (requestError) {
      updateError(requestError, '批量操作失败');
    } finally {
      setBusyAction('');
    }
  }

  async function handleDeleteActiveMessage() {
    if (!activeMailbox || !activeMessage) {
      return;
    }
    setBusyAction('delete');
    try {
      await deleteMessage({
        mailboxId: activeMailbox.id,
        method: activeMethod,
        folder: activeMessage.folderId,
        messageId: activeMessage.id,
      });
      setActiveMessage(null);
      setActiveMessageId('');
      updateNotice('邮件已删除');
      await loadMessagesForFolder(page);
      await loadOpsForMailbox(activeMailbox.id);
    } catch (requestError) {
      updateError(requestError, '删除邮件失败');
    } finally {
      setBusyAction('');
    }
  }

  async function handleDownloadAttachment(attachmentId: string, filename: string) {
    if (!activeMailbox || !activeMessage) {
      return;
    }
    setBusyAction(`download-${attachmentId}`);
    try {
      const attachment = await downloadAttachment({
        mailboxId: activeMailbox.id,
        method: activeMethod,
        folder: activeMessage.folderId,
        messageId: activeMessage.id,
        attachmentId,
      });
      triggerFileDownload(
        attachment.name || filename,
        base64ToBlob(attachment.content_base64 ?? '', attachment.content_type || 'application/octet-stream'),
      );
      updateNotice('附件已下载');
    } catch (requestError) {
      updateError(requestError, '下载附件失败');
    } finally {
      setBusyAction('');
    }
  }

  function openCompose(mode: ComposeMode, message?: Email | null) {
    if (!message) {
      setCompose({ ...EMPTY_COMPOSE, open: true, mode: 'new' });
      return;
    }

    const activeEmail = activeMailbox?.email ?? '';
    const replyAllRecipients = Array.from(
      new Set([message.sender, ...message.to_recipients, ...message.cc_recipients].filter((item) => item && item !== activeEmail)),
    );

    setCompose({
      open: true,
      mode,
      messageId: message.id,
      draftMessageId: '',
      subject:
        mode === 'forward'
          ? titleWithPrefix('Fwd: ', message.subject)
          : titleWithPrefix('Re: ', message.subject),
      to: mode === 'reply' ? message.sender : mode === 'replyAll' ? joinRecipients(replyAllRecipients) : '',
      cc: '',
      bcc: '',
      bodyText: makeQuotedBody(message),
      attachments: [],
      attachmentNames: [],
      submitting: false,
    });
  }

  function openDraftEditor(message: Email) {
    setCompose({
      open: true,
      mode: 'new',
      messageId: message.id,
      draftMessageId: message.id,
      subject: message.subject,
      to: joinRecipients(message.to_recipients),
      cc: joinRecipients(message.cc_recipients),
      bcc: joinRecipients(message.bcc_recipients),
      bodyText: message.body_text,
      attachments: [],
      attachmentNames: [],
      submitting: false,
    });
  }

  async function handleComposeFiles(files: FileList | null) {
    if (!files || files.length === 0) {
      return;
    }
    try {
      const uploaded = await Promise.all(Array.from(files).map((file) => fileToAttachmentPayload(file)));
      setCompose((current) => ({
        ...current,
        attachments: [...current.attachments, ...uploaded],
        attachmentNames: [...current.attachmentNames, ...uploaded.map((item) => item.name)],
      }));
    } catch (requestError) {
      updateError(requestError, '读取附件失败');
    }
  }

  async function submitCompose(sendNow: boolean) {
    if (!activeMailbox) {
      return;
    }
    setCompose((current) => ({ ...current, submitting: true }));
    try {
      const payload = {
        mailboxId: activeMailbox.id,
        method: activeMethod,
        messageId: compose.mode === 'new' ? undefined : compose.messageId,
        draftMessageId: compose.draftMessageId || undefined,
        subject: compose.subject,
        bodyText: compose.bodyText,
        toRecipients: splitRecipients(compose.to),
        ccRecipients: splitRecipients(compose.cc),
        bccRecipients: splitRecipients(compose.bcc),
        attachments: compose.attachments,
        sendNow,
      };

      const result =
        compose.mode === 'new'
          ? sendNow
            ? await sendMessage(payload)
            : await saveDraft(payload)
          : compose.mode === 'reply'
            ? await replyMessage(payload)
            : compose.mode === 'replyAll'
              ? await replyAllMessage(payload)
              : await forwardMessage(payload);

      if (sendNow) {
        setCompose(EMPTY_COMPOSE);
        updateNotice('邮件已发送');
      } else {
        setCompose((current) => ({
          ...current,
          draftMessageId: result.id,
          messageId: result.id,
          submitting: false,
        }));
        updateNotice('草稿已保存');
      }

      if (activeMailbox) {
        await loadOpsForMailbox(activeMailbox.id);
        if (activeFolder?.type === 'drafts' || activeFolder?.type === 'sent') {
          await loadMessagesForFolder(page);
        }
      }
    } catch (requestError) {
      updateError(requestError, sendNow ? '发送邮件失败' : '保存草稿失败');
    } finally {
      setCompose((current) => ({ ...current, submitting: false }));
    }
  }

  async function handleSaveMeta() {
    if (!activeMailbox || !activeMessage) {
      return;
    }
    setBusyAction('meta');
    try {
      const payload = await updateMessageMeta({
        mailboxId: activeMailbox.id,
        method: activeMethod,
        folder: activeMessage.folderId,
        messageId: activeMessage.id,
        tags: splitRecipients(metaForm.tags),
        followUp: metaForm.followUp,
        notes: metaForm.notes,
        snoozedUntil: fromDatetimeLocalValue(metaForm.snoozedUntil),
        status: metaForm.status,
      });
      const nextMessage: Email = {
        ...activeMessage,
        meta: payload.meta as MessageMeta,
      };
      if (payload.message) {
        updateMessageCollection(payload.message);
      } else {
        updateMessageCollection(nextMessage);
      }
      updateNotice('标签、备注和提醒已更新');
      await loadOpsForMailbox(activeMailbox.id);
    } catch (requestError) {
      updateError(requestError, '更新消息元数据失败');
    } finally {
      setBusyAction('');
    }
  }

  async function handleFolderAction(type: 'create' | 'rename' | 'delete') {
    if (!activeMailbox) {
      return;
    }
    try {
      if (type === 'create') {
        const name = window.prompt('请输入新文件夹名称');
        if (!name) {
          return;
        }
        await createFolder({ mailboxId: activeMailbox.id, method: activeMethod, displayName: name });
        updateNotice('文件夹已创建');
      }
      if (type === 'rename' && activeFolder) {
        const name = window.prompt('请输入新的文件夹名称', activeFolder.displayName);
        if (!name) {
          return;
        }
        await renameFolder({
          mailboxId: activeMailbox.id,
          method: activeMethod,
          folderId: activeFolder.id,
          displayName: name,
        });
        updateNotice('文件夹已重命名');
      }
      if (type === 'delete' && activeFolder) {
        if (!window.confirm(`确认删除文件夹 ${activeFolder.displayName} 吗？`)) {
          return;
        }
        await deleteFolder({ mailboxId: activeMailbox.id, method: activeMethod, folderId: activeFolder.id });
        updateNotice('文件夹已删除');
      }
      await loadFoldersForMailbox(activeMailbox.id, activeMethod);
    } catch (requestError) {
      updateError(requestError, '文件夹操作失败');
    }
  }

  async function executeGlobalSearch(nextPage = 1) {
    if (!globalSearchQuery.trim()) {
      setGlobalSearchResults([]);
      setGlobalSearchMeta(EMPTY_PAGINATION);
      return;
    }
    setGlobalSearchLoading(true);
    try {
      const payload = await searchMessages({
        query: globalSearchQuery.trim(),
        mailboxIds: !searchAcrossAll && activeMailbox ? [activeMailbox.id] : undefined,
        page: nextPage,
        pageSize: 20,
        sortOrder: 'desc',
      });
      setGlobalSearchResults(payload.items);
      setGlobalSearchMeta(payload.meta);
    } catch (requestError) {
      updateError(requestError, '统一搜索失败');
    } finally {
      setGlobalSearchLoading(false);
    }
  }

  function jumpToMessage(message: Email) {
    setSearchOpen(false);
    setSearchQuery('');
    setFilterUnread(false);
    setFilterStarred(false);
    setFilterAttachment(false);
    setPage(1);
    const targetMailboxId = message.mailboxId || activeMailbox?.id || '';
    if (!targetMailboxId) {
      return;
    }
    setActiveMailboxId(targetMailboxId);
    setPendingFocus({
      mailboxId: targetMailboxId,
      folderId: message.folderId || 'INBOX',
      messageId: message.id,
    });
  }

  async function handleRunSync() {
    if (!activeMailbox) {
      return;
    }
    setBusyAction('sync');
    try {
      await runSync({
        mailboxId: activeMailbox.id,
        method: activeMethod,
        folderLimit: 5,
        messageLimit: 20,
        includeBody: true,
        applyRules: true,
      });
      updateNotice('同步任务已执行');
      await Promise.all([loadOpsForMailbox(activeMailbox.id), loadMessagesForFolder(page)]);
    } catch (requestError) {
      updateError(requestError, '执行同步失败');
    } finally {
      setBusyAction('');
    }
  }

  function openCreateRule() {
    setRuleEditor(EMPTY_RULE_EDITOR);
  }

  function openEditRule(rule: RuleRecord) {
    setRuleEditor({
      id: rule.id,
      name: rule.name,
      enabled: rule.enabled,
      priority: rule.priority,
      conditionsText: JSON.stringify(rule.conditions, null, 2),
      actionsText: JSON.stringify(rule.actions, null, 2),
    });
  }

  async function saveRuleEditor() {
    if (!activeMailbox || !ruleEditor.name.trim()) {
      return;
    }
    setBusyAction('rule-save');
    try {
      const conditions = parseJsonObject(ruleEditor.conditionsText, '条件');
      const actions = parseJsonObject(ruleEditor.actionsText, '动作');
      if (ruleEditor.id) {
        await updateRule({
          mailboxId: activeMailbox.id,
          ruleId: ruleEditor.id,
          name: ruleEditor.name.trim(),
          enabled: ruleEditor.enabled,
          priority: ruleEditor.priority,
          conditions,
          actions,
        });
        updateNotice('规则已更新');
      } else {
        await createRule({
          mailboxId: activeMailbox.id,
          name: ruleEditor.name.trim(),
          enabled: ruleEditor.enabled,
          priority: ruleEditor.priority,
          conditions,
          actions,
        });
        updateNotice('规则已创建');
      }
      setRuleEditor(EMPTY_RULE_EDITOR);
      await Promise.all([loadRulesForMailbox(activeMailbox.id), loadOpsForMailbox(activeMailbox.id)]);
    } catch (requestError) {
      updateError(requestError, '保存规则失败');
    } finally {
      setBusyAction('');
    }
  }

  async function handleDeleteRule(ruleId: number) {
    if (!activeMailbox || !window.confirm('确认删除这条规则吗？')) {
      return;
    }
    setBusyAction('rule-delete');
    try {
      await deleteRule({ mailboxId: activeMailbox.id, ruleId });
      updateNotice('规则已删除');
      await Promise.all([loadRulesForMailbox(activeMailbox.id), loadOpsForMailbox(activeMailbox.id)]);
    } catch (requestError) {
      updateError(requestError, '删除规则失败');
    } finally {
      setBusyAction('');
    }
  }

  async function handleApplyRules(ruleId?: number) {
    if (!activeMailbox) {
      return;
    }
    setBusyAction('rule-apply');
    try {
      const payload = await applyRules({
        mailboxId: activeMailbox.id,
        method: activeMethod,
        folder: activeFolderId || undefined,
        ruleId,
        limit: 200,
      });
      updateNotice(`规则执行完成，共 ${payload.count} 条`);
      await Promise.all([loadMessagesForFolder(page), loadOpsForMailbox(activeMailbox.id)]);
    } catch (requestError) {
      updateError(requestError, '应用规则失败');
    } finally {
      setBusyAction('');
    }
  }

  return (
    <div className="flex h-full w-full overflow-hidden bg-white dark:bg-slate-950">
      <div className="hidden w-72 shrink-0 overflow-hidden border-r border-slate-200 bg-slate-50/80 dark:border-slate-800 dark:bg-slate-900/50 lg:flex lg:flex-col">
        <div
          className={cn(
            'flex flex-col overflow-hidden p-4',
            foldersCollapsed
              ? 'min-h-0 flex-1'
              : 'max-h-[45%] min-h-[11rem] shrink-0 border-b border-slate-200 dark:border-slate-800',
          )}
        >
          <div className="mb-2 flex items-center justify-between">
            <h2 className="text-sm font-semibold uppercase tracking-[0.24em] text-slate-500">邮箱档案</h2>
            <Badge variant="outline">{accounts.length}</Badge>
          </div>
          <div className="min-h-0 flex-1 space-y-1 overflow-y-auto pr-1">
            {loadingAccounts ? (
              <div className="flex items-center gap-2 rounded-md px-2 py-2 text-sm text-slate-500">
                <LoaderCircle className="h-4 w-4 animate-spin" />
                {t('loading')}
              </div>
            ) : accounts.length === 0 ? (
              <div className="rounded-md border border-dashed border-slate-200 px-3 py-4 text-sm text-slate-500 dark:border-slate-800">
                暂无邮箱档案
              </div>
            ) : (
              accounts.map((account) => (
                <button
                  key={account.id}
                  onClick={() => {
                    setActiveMailboxId(account.id);
                    setPage(1);
                    setSelectedMessageIds([]);
                    setActiveMessageId('');
                    setActiveMessage(null);
                  }}
                  className={cn(
                    'w-full rounded-lg border px-3 py-2 text-left transition-colors',
                    activeMailboxId === account.id
                      ? 'border-slate-900 bg-white text-slate-900 shadow-sm dark:border-slate-200 dark:bg-slate-950 dark:text-slate-50'
                      : 'border-transparent hover:bg-white dark:hover:bg-slate-950',
                  )}
                >
                  <div className="flex items-center justify-between gap-2">
                    <div className="min-w-0">
                      <div className="truncate text-sm font-medium">{account.label || account.email}</div>
                      <div className="truncate text-xs text-slate-500">{account.email}</div>
                    </div>
                    <div
                      className={cn(
                        'h-2.5 w-2.5 shrink-0 rounded-full',
                        account.status === 'connected'
                          ? 'bg-emerald-500'
                          : account.status === 'disconnected'
                            ? 'bg-red-500'
                            : 'bg-slate-300',
                      )}
                    />
                  </div>
                  <div className="mt-2 text-[11px] text-slate-500">{methodLabel(account.preferredMethod)}</div>
                </button>
              ))
            )}
          </div>
        </div>

        <div
          className={cn(
            'flex min-h-0 flex-col overflow-hidden p-4',
            foldersCollapsed ? 'shrink-0 border-t border-slate-200 dark:border-slate-800' : 'flex-1',
          )}
        >
          <div className={cn('flex items-center justify-between gap-2', foldersCollapsed ? 'mb-0' : 'mb-3')}>
            <button
              type="button"
              className="flex min-w-0 items-center gap-2 rounded-md text-left transition-colors hover:text-slate-900 dark:hover:text-slate-50"
              onClick={() => setFoldersCollapsed((current) => !current)}
              aria-expanded={!foldersCollapsed}
              aria-controls="workspace-folders-panel"
            >
              <ChevronRight
                className={cn(
                  'h-4 w-4 shrink-0 text-slate-400 transition-transform duration-200',
                  !foldersCollapsed && 'rotate-90',
                )}
              />
              <h2 className="text-sm font-semibold uppercase tracking-[0.24em] text-slate-500">{t('folders')}</h2>
            </button>
            <div className="flex items-center gap-1">
              <Button variant="ghost" size="icon" className="h-7 w-7" onClick={() => void handleFolderAction('create')}>
                <FolderPlus className="h-4 w-4" />
              </Button>
              <Button
                variant="ghost"
                size="icon"
                className="h-7 w-7"
                disabled={!activeFolder || activeFolder.type !== 'custom'}
                onClick={() => void handleFolderAction('rename')}
              >
                <FolderPen className="h-4 w-4" />
              </Button>
              <Button
                variant="ghost"
                size="icon"
                className="h-7 w-7"
                disabled={!activeFolder || activeFolder.type !== 'custom'}
                onClick={() => void handleFolderAction('delete')}
              >
                <Trash2 className="h-4 w-4" />
              </Button>
            </div>
          </div>

          <div
            id="workspace-folders-panel"
            className={cn(
              'grid transition-[grid-template-rows,opacity] duration-200 ease-out',
              foldersCollapsed ? 'grid-rows-[0fr] opacity-0' : 'min-h-0 flex-1 grid-rows-[1fr] opacity-100',
            )}
          >
            <div className="min-h-0 overflow-hidden">
              <div className="min-h-0 h-full space-y-1 overflow-y-auto pr-1">
                {loadingFolders ? (
                  <div className="flex items-center gap-2 rounded-md px-2 py-2 text-sm text-slate-500">
                    <LoaderCircle className="h-4 w-4 animate-spin" />
                    {t('loading')}
                  </div>
                ) : (
                  folders.map((folder) => (
                    <button
                      key={folder.id}
                      onClick={() => {
                        setActiveFolderId(folder.id);
                        setPage(1);
                        setSelectedMessageIds([]);
                      }}
                      className={cn(
                        'flex w-full items-center justify-between rounded-md px-2 py-2 text-left text-sm transition-colors',
                        activeFolderId === folder.id
                          ? 'bg-slate-200 font-medium text-slate-900 dark:bg-slate-800 dark:text-slate-50'
                          : 'text-slate-600 hover:bg-slate-200/70 dark:text-slate-300 dark:hover:bg-slate-800/70',
                      )}
                    >
                      <span className="truncate">{folder.displayName}</span>
                      {folder.unreadCount > 0 ? <Badge variant="secondary">{folder.unreadCount}</Badge> : null}
                    </button>
                  ))
                )}
              </div>
            </div>
          </div>
        </div>
      </div>

      <div
        className={cn(
          'flex w-full shrink-0 flex-col border-r border-slate-200 bg-white dark:border-slate-800 dark:bg-slate-950 md:w-[28rem]',
          activeMessageId ? 'hidden md:flex' : 'flex',
        )}
      >
        <div className="space-y-3 border-b border-slate-200 p-3 dark:border-slate-800">
          <div className="flex flex-wrap items-center gap-2">
            <div className="relative min-w-[220px] flex-1">
              <Search className="absolute left-2.5 top-2 h-4 w-4 text-slate-400" />
              <Input
                value={searchQuery}
                onChange={(event) => {
                  setSearchQuery(event.target.value);
                  setPage(1);
                }}
                placeholder={t('search')}
                className="pl-8"
              />
            </div>
            <Button variant="outline" onClick={() => void refreshEverything()} disabled={loadingMessages || busyAction !== ''}>
              <RefreshCw className="mr-2 h-4 w-4" />
              {t('refresh')}
            </Button>
            <Button onClick={() => openCompose('new')}>
              <FilePlus2 className="mr-2 h-4 w-4" />
              {t('compose')}
            </Button>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <Button variant={filterUnread ? 'secondary' : 'ghost'} size="sm" onClick={() => setFilterUnread((current) => !current)}>
              {t('unread')}
            </Button>
            <Button variant={filterStarred ? 'secondary' : 'ghost'} size="sm" onClick={() => setFilterStarred((current) => !current)}>
              {t('starred')}
            </Button>
            <Button
              variant={filterAttachment ? 'secondary' : 'ghost'}
              size="sm"
              onClick={() => setFilterAttachment((current) => !current)}
            >
              {t('hasAttachment')}
            </Button>
            <Button variant="ghost" size="sm" onClick={() => setSortOrder((current) => (current === 'desc' ? 'asc' : 'desc'))}>
              {sortOrder === 'desc' ? t('dateDesc') : t('dateAsc')}
            </Button>
            <Button variant="ghost" size="sm" onClick={() => setSearchOpen(true)}>
              统一搜索
            </Button>
            <Button variant="ghost" size="sm" onClick={() => setRulesOpen(true)}>
              规则
            </Button>
            <Button variant="ghost" size="sm" onClick={() => setOpsOpen(true)}>
              同步中心
            </Button>
          </div>

          {selectedMessageIds.length > 0 ? (
            <div className="flex flex-wrap items-center gap-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-sm dark:border-slate-800 dark:bg-slate-900/50">
              <span className="mr-2 font-medium">{selectedMessageIds.length} 已选</span>
              <Button variant="secondary" size="sm" onClick={() => setSelectedMessageIds([])}>
                <X className="mr-2 h-4 w-4" />
                {t('clearSelection')}
              </Button>
              <Button variant="secondary" size="sm" onClick={() => void handleBatchAction('mark_read', selectedMessageIds)}>
                {t('markRead')}
              </Button>
              <Button variant="secondary" size="sm" onClick={() => void handleBatchAction('mark_unread', selectedMessageIds)}>
                {t('markUnread')}
              </Button>
              <Button variant="secondary" size="sm" onClick={() => void handleBatchAction('flag', selectedMessageIds)}>
                {t('star')}
              </Button>
              <Button variant="secondary" size="sm" onClick={() => void handleBatchAction('archive', selectedMessageIds)}>
                {t('archive')}
              </Button>
              <Button
                variant="secondary"
                size="sm"
                onClick={() => {
                  const destinationFolder = window.prompt('请输入目标文件夹 ID');
                  if (destinationFolder) {
                    void handleBatchAction('move', selectedMessageIds, destinationFolder);
                  }
                }}
              >
                {t('move')}
              </Button>
              <Button variant="destructive" size="sm" onClick={() => void handleBatchAction('delete', selectedMessageIds)}>
                {t('delete')}
              </Button>
            </div>
          ) : null}

          {error ? <div className="rounded-md border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-600">{error}</div> : null}
          {notice ? (
            <div className="rounded-md border border-emerald-200 bg-emerald-50 px-3 py-2 text-sm text-emerald-700">{notice}</div>
          ) : null}
        </div>

        <div className="min-h-0 flex-1 overflow-y-auto">
          {loadingMessages ? (
            <div className="flex h-full items-center justify-center gap-2 text-sm text-slate-500">
              <LoaderCircle className="h-4 w-4 animate-spin" />
              {t('loading')}
            </div>
          ) : messages.length === 0 ? (
            <div className="flex h-full items-center justify-center text-sm text-slate-500">{t('noEmails')}</div>
          ) : (
            <div className="divide-y divide-slate-100 dark:divide-slate-800/70">
              {messages.map((message) => (
                <div
                  key={message.id}
                  onClick={() => void loadMessageDetail(message.id, message.folderId)}
                  className={cn(
                    'cursor-pointer px-3 py-3 transition-colors hover:bg-slate-50 dark:hover:bg-slate-900/50',
                    activeMessageId === message.id && 'bg-slate-100 dark:bg-slate-800/70',
                    !message.is_read && 'font-medium',
                  )}
                >
                  <div className="flex items-start gap-3">
                    <input
                      type="checkbox"
                      className="mt-1 rounded border-slate-300"
                      checked={selectedMessageIds.includes(message.id)}
                      onChange={(event) => {
                        event.stopPropagation();
                        setSelectedMessageIds((current) =>
                          current.includes(message.id) ? current.filter((item) => item !== message.id) : [...current, message.id],
                        );
                      }}
                    />
                    <div className="min-w-0 flex-1">
                      <div className="mb-1 flex items-center justify-between gap-3">
                        <span className="truncate text-sm text-slate-900 dark:text-slate-100">{message.sender}</span>
                        <span className="shrink-0 text-[11px] text-slate-500">
                          {message.date ? format(new Date(message.date), 'MM-dd HH:mm') : '-'}
                        </span>
                      </div>
                      <div className="truncate text-sm text-slate-800 dark:text-slate-200">{message.subject}</div>
                      <div className="mt-1 line-clamp-2 text-xs text-slate-500 dark:text-slate-400">{message.preview}</div>
                      <div className="mt-2 flex items-center gap-2">
                        {message.is_flagged ? <Star className="h-3.5 w-3.5 fill-amber-400 text-amber-400" /> : null}
                        {message.has_attachments ? <Paperclip className="h-3.5 w-3.5 text-slate-400" /> : null}
                        {message.importance === 'high' ? <AlertCircle className="h-3.5 w-3.5 text-red-500" /> : null}
                        {message.meta.tags.length > 0 ? (
                          <Badge variant="outline" className="max-w-[9rem] truncate">
                            {message.meta.tags.join(', ')}
                          </Badge>
                        ) : null}
                      </div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        <div className="flex items-center justify-between border-t border-slate-200 px-3 py-2 text-xs text-slate-500 dark:border-slate-800">
          <label className="flex items-center gap-2">
            <input
              type="checkbox"
              className="rounded border-slate-300"
              checked={allVisibleSelected}
              onChange={(event) => setSelectedMessageIds(event.target.checked ? messages.map((item) => item.id) : [])}
            />
            {t('selectAll')}
          </label>
          <div className="flex items-center gap-2">
            <span>
              {t('page')} {listMeta.page || 1} {t('of')} {listMeta.total_pages || 1}
            </span>
            <Button variant="ghost" size="icon" className="h-7 w-7" disabled={!listMeta.has_prev} onClick={() => setPage((current) => Math.max(1, current - 1))}>
              <ChevronLeft className="h-4 w-4" />
            </Button>
            <Button
              variant="ghost"
              size="icon"
              className="h-7 w-7"
              disabled={!listMeta.has_next}
              onClick={() => setPage((current) => current + 1)}
            >
              <ChevronRight className="h-4 w-4" />
            </Button>
          </div>
        </div>
      </div>

      <div className={cn('min-w-0 flex-1 bg-white dark:bg-slate-950', !activeMessageId ? 'hidden md:flex md:flex-col' : 'flex flex-col')}>
        {loadingDetail ? (
          <div className="flex flex-1 items-center justify-center gap-2 text-sm text-slate-500">
            <LoaderCircle className="h-4 w-4 animate-spin" />
            {t('loading')}
          </div>
        ) : activeMessage ? (
          <>
            <div className="flex shrink-0 flex-wrap items-center justify-between gap-3 border-b border-slate-200 px-4 py-3 dark:border-slate-800">
              <div className="flex flex-wrap items-center gap-2">
                <Button variant="ghost" size="icon" className="md:hidden" onClick={() => setActiveMessageId('')}>
                  <ChevronLeft className="h-4 w-4" />
                </Button>
                {isDraftFolder(activeMessage, activeFolderId) ? (
                  <Button variant="outline" size="sm" onClick={() => openDraftEditor(activeMessage)}>
                    编辑草稿
                  </Button>
                ) : null}
                <Button variant="ghost" size="sm" onClick={() => openCompose('reply', activeMessage)}>
                  <Reply className="mr-2 h-4 w-4" />
                  回复
                </Button>
                <Button variant="ghost" size="sm" onClick={() => openCompose('replyAll', activeMessage)}>
                  <ReplyAll className="mr-2 h-4 w-4" />
                  回复全部
                </Button>
                <Button variant="ghost" size="sm" onClick={() => openCompose('forward', activeMessage)}>
                  <Forward className="mr-2 h-4 w-4" />
                  转发
                </Button>
                <Button variant="ghost" size="sm" onClick={() => void handleSingleStateUpdate('read', !activeMessage.is_read)}>
                  {activeMessage.is_read ? <Mail className="mr-2 h-4 w-4" /> : <MailOpen className="mr-2 h-4 w-4" />}
                  {activeMessage.is_read ? t('markUnread') : t('markRead')}
                </Button>
                <Button variant="ghost" size="sm" onClick={() => void handleSingleStateUpdate('flag', !activeMessage.is_flagged)}>
                  <Star className={cn('mr-2 h-4 w-4', activeMessage.is_flagged && 'fill-amber-400 text-amber-400')} />
                  {activeMessage.is_flagged ? t('unstar') : t('star')}
                </Button>
                <Button variant="ghost" size="sm" onClick={() => void handleBatchAction('archive', [activeMessage.id])}>
                  <Archive className="mr-2 h-4 w-4" />
                  {t('archive')}
                </Button>
                <Button variant="destructive" size="sm" onClick={() => void handleDeleteActiveMessage()}>
                  <Trash2 className="mr-2 h-4 w-4" />
                  {t('delete')}
                </Button>
              </div>
              <div className="text-xs text-slate-500">{methodLabel(activeMethod)}</div>
            </div>

            <div className="min-h-0 flex-1 overflow-y-auto px-6 py-5">
              <div className="mb-6">
                <div className="mb-2 flex items-start justify-between gap-4">
                  <h1 className="text-2xl font-semibold text-slate-900 dark:text-slate-50">{activeMessage.subject}</h1>
                  <div className="flex items-center gap-1 text-sm text-slate-500">
                    <Clock className="h-3.5 w-3.5" />
                    {activeMessage.date ? format(new Date(activeMessage.date), 'yyyy-MM-dd HH:mm') : '-'}
                  </div>
                </div>
                <div className="space-y-1 text-sm text-slate-600 dark:text-slate-300">
                  <div>
                    <span className="font-medium">{activeMessage.sender}</span>
                    <span className="ml-2 text-slate-400">{activeMessage.mailboxEmail || activeMailbox?.email}</span>
                  </div>
                  <div>{t('to')}: {activeMessage.to_recipients.join(', ') || '-'}</div>
                  {activeMessage.cc_recipients.length > 0 ? <div>{t('cc')}: {activeMessage.cc_recipients.join(', ')}</div> : null}
                </div>
              </div>

              <div className={cn('mb-6 grid gap-4', !detailSidebarCollapsed && 'xl:grid-cols-[minmax(0,1.3fr)_minmax(320px,0.9fr)]')}>
                <div className="rounded-xl border border-slate-200 p-4 dark:border-slate-800">
                  <div className="mb-3 flex items-center justify-between">
                    <h3 className="text-sm font-semibold">消息视图</h3>
                    <div className="flex items-center gap-1">
                      <Button variant={viewMode === 'html' ? 'secondary' : 'ghost'} size="sm" onClick={() => setViewMode('html')}>
                        {t('htmlView')}
                      </Button>
                      <Button variant={viewMode === 'text' ? 'secondary' : 'ghost'} size="sm" onClick={() => setViewMode('text')}>
                        {t('textView')}
                      </Button>
                      <Button variant={viewMode === 'headers' ? 'secondary' : 'ghost'} size="sm" onClick={() => setViewMode('headers')}>
                        {t('headersView')}
                      </Button>
                      <Button
                        variant={detailSidebarCollapsed ? 'ghost' : 'secondary'}
                        size="sm"
                        onClick={() => setDetailSidebarCollapsed((current) => !current)}
                        title={detailSidebarCollapsed ? '展开侧边信息' : '收起侧边信息'}
                      >
                        {detailSidebarCollapsed ? <ChevronLeft className="mr-2 h-4 w-4" /> : <ChevronRight className="mr-2 h-4 w-4" />}
                        侧边信息
                      </Button>
                    </div>
                  </div>

                  {viewMode === 'html' ? (
                    activeMessage.body_html ? (
                      <div className="prose prose-sm max-w-none dark:prose-invert" dangerouslySetInnerHTML={{ __html: activeMessage.body_html }} />
                    ) : (
                      <pre className="whitespace-pre-wrap text-sm text-slate-700 dark:text-slate-300">{activeMessage.body_text}</pre>
                    )
                  ) : null}
                  {viewMode === 'text' ? (
                    <pre className="whitespace-pre-wrap text-sm text-slate-700 dark:text-slate-300">{activeMessage.body_text}</pre>
                  ) : null}
                  {viewMode === 'headers' ? (
                    <pre className="whitespace-pre-wrap rounded-lg bg-slate-50 p-3 text-xs text-slate-600 dark:bg-slate-900 dark:text-slate-300">
                      {activeMessage.headers || '(无 Headers)'}
                    </pre>
                  ) : null}

                  {activeMessage.attachments.length > 0 ? (
                    <div className="mt-6 border-t border-slate-200 pt-4 dark:border-slate-800">
                      <h3 className="mb-3 flex items-center gap-2 text-sm font-semibold">
                        <Paperclip className="h-4 w-4" />
                        {t('attachments')} ({activeMessage.attachments.length})
                      </h3>
                      <div className="space-y-2">
                        {activeMessage.attachments.map((attachment) => (
                          <div
                            key={attachment.id}
                            className="flex items-center justify-between gap-3 rounded-lg border border-slate-200 px-3 py-2 dark:border-slate-800"
                          >
                            <div className="min-w-0">
                              <div className="truncate text-sm font-medium">{attachment.name}</div>
                              <div className="text-xs text-slate-500">{(attachment.size / 1024).toFixed(1)} KB</div>
                            </div>
                            <Button
                              variant="ghost"
                              size="icon"
                              className="h-8 w-8"
                              disabled={busyAction === `download-${attachment.id}`}
                              onClick={() => void handleDownloadAttachment(attachment.id, attachment.name)}
                            >
                              <Download className="h-4 w-4" />
                            </Button>
                          </div>
                        ))}
                      </div>
                    </div>
                  ) : null}
                </div>

                {!detailSidebarCollapsed ? (
                <div className="space-y-4">
                  <div className="rounded-xl border border-slate-200 p-4 dark:border-slate-800">
                    <div className={cn('flex items-center justify-between gap-2', metaPanelCollapsed ? 'mb-0' : 'mb-3')}>
                      <button
                        type="button"
                        className="flex min-w-0 items-center gap-2 rounded-md text-left transition-colors hover:text-slate-900 dark:hover:text-slate-50"
                        onClick={() => setMetaPanelCollapsed((current) => !current)}
                        aria-expanded={!metaPanelCollapsed}
                        aria-controls="workspace-meta-panel"
                      >
                        <ChevronRight
                          className={cn(
                            'h-4 w-4 shrink-0 text-slate-400 transition-transform duration-200',
                            !metaPanelCollapsed && 'rotate-90',
                          )}
                        />
                        <h3 className="text-sm font-semibold">标签 / 跟进 / 备注 / Snooze</h3>
                      </button>
                      <Badge variant="outline">{activeMessage.meta.status || 'active'}</Badge>
                    </div>
                    <div
                      id="workspace-meta-panel"
                      className={cn(
                        'grid transition-[grid-template-rows,opacity] duration-200 ease-out',
                        metaPanelCollapsed ? 'grid-rows-[0fr] opacity-0' : 'grid-rows-[1fr] opacity-100',
                      )}
                    >
                      <div className="min-h-0 overflow-hidden">
                        <div className="grid gap-3 pt-1">
                          <div>
                            <label className="mb-1 block text-xs font-medium text-slate-500">标签</label>
                            <Input value={metaForm.tags} onChange={(event) => setMetaForm((current) => ({ ...current, tags: event.target.value }))} placeholder="vip, follow-up" />
                          </div>
                          <div>
                            <label className="mb-1 block text-xs font-medium text-slate-500">跟进</label>
                            <Input value={metaForm.followUp} onChange={(event) => setMetaForm((current) => ({ ...current, followUp: event.target.value }))} placeholder="today / tomorrow / custom" />
                          </div>
                          <div>
                            <label className="mb-1 block text-xs font-medium text-slate-500">稍后提醒</label>
                            <Input type="datetime-local" value={metaForm.snoozedUntil} onChange={(event) => setMetaForm((current) => ({ ...current, snoozedUntil: event.target.value }))} />
                          </div>
                          <div>
                            <label className="mb-1 block text-xs font-medium text-slate-500">状态</label>
                            <select
                              className="flex h-9 w-full rounded-md border border-slate-200 bg-transparent px-3 text-sm dark:border-slate-800"
                              value={metaForm.status}
                              onChange={(event) => setMetaForm((current) => ({ ...current, status: event.target.value }))}
                            >
                              <option value="active">active</option>
                              <option value="snoozed">snoozed</option>
                              <option value="done">done</option>
                            </select>
                          </div>
                          <div>
                            <label className="mb-1 block text-xs font-medium text-slate-500">备注</label>
                            <textarea
                              className="min-h-28 w-full rounded-md border border-slate-200 bg-transparent px-3 py-2 text-sm shadow-sm dark:border-slate-800"
                              value={metaForm.notes}
                              onChange={(event) => setMetaForm((current) => ({ ...current, notes: event.target.value }))}
                            />
                          </div>
                          <Button onClick={() => void handleSaveMeta()} disabled={busyAction === 'meta'}>
                            {busyAction === 'meta' ? <LoaderCircle className="mr-2 h-4 w-4 animate-spin" /> : null}
                            保存 Meta
                          </Button>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="rounded-xl border border-slate-200 p-4 dark:border-slate-800">
                    <div className={cn('flex items-center justify-between gap-2', threadPanelCollapsed ? 'mb-0' : 'mb-3')}>
                      <button
                        type="button"
                        className="flex min-w-0 items-center gap-2 rounded-md text-left transition-colors hover:text-slate-900 dark:hover:text-slate-50"
                        onClick={() => setThreadPanelCollapsed((current) => !current)}
                        aria-expanded={!threadPanelCollapsed}
                        aria-controls="workspace-thread-panel"
                      >
                        <ChevronRight
                          className={cn(
                            'h-4 w-4 shrink-0 text-slate-400 transition-transform duration-200',
                            !threadPanelCollapsed && 'rotate-90',
                          )}
                        />
                        <h3 className="text-sm font-semibold">线程聚合视图</h3>
                      </button>
                      <Badge variant="outline">{threadItems.length}</Badge>
                    </div>
                    <div
                      id="workspace-thread-panel"
                      className={cn(
                        'grid transition-[grid-template-rows,opacity] duration-200 ease-out',
                        threadPanelCollapsed ? 'grid-rows-[0fr] opacity-0' : 'grid-rows-[1fr] opacity-100',
                      )}
                    >
                      <div className="min-h-0 overflow-hidden">
                        <div className="pt-1">
                          {threadItems.length === 0 ? (
                            <div className="text-sm text-slate-500">当前会话没有已缓存的线程邮件</div>
                          ) : (
                            <div className="space-y-2">
                              {threadItems.map((item) => (
                                <button
                                  key={item.id}
                                  onClick={() => void loadMessageDetail(item.id, item.folderId)}
                                  className={cn(
                                    'w-full rounded-lg border px-3 py-2 text-left transition-colors',
                                    item.id === activeMessage.id
                                      ? 'border-slate-900 bg-slate-50 dark:border-slate-200 dark:bg-slate-900'
                                      : 'border-slate-200 hover:bg-slate-50 dark:border-slate-800 dark:hover:bg-slate-900/60',
                                  )}
                                >
                                  <div className="truncate text-sm font-medium">{item.subject}</div>
                                  <div className="mt-1 text-xs text-slate-500">
                                    {item.sender} · {item.date ? format(new Date(item.date), 'MM-dd HH:mm') : '-'}
                                  </div>
                                </button>
                              ))}
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
                ) : null}
              </div>

              <div className="rounded-xl border border-slate-200 p-4 text-xs text-slate-500 dark:border-slate-800">
                <div className="flex gap-2">
                  <span className="min-w-24 font-medium">{t('messageId')}:</span>
                  <span className="break-all font-mono">{activeMessage.internet_message_id || activeMessage.id}</span>
                </div>
                <div className="mt-2 flex gap-2">
                  <span className="min-w-24 font-medium">{t('conversationId')}:</span>
                  <span className="break-all font-mono">{activeMessage.conversation_id || '-'}</span>
                </div>
              </div>
            </div>
          </>
        ) : (
          <div className="flex flex-1 flex-col items-center justify-center gap-4 text-slate-500">
            <Mail className="h-12 w-12 opacity-20" />
            <p>{t('selectEmail')}</p>
          </div>
        )}
      </div>

      {compose.open ? (
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/45 px-4 py-6">
          <div className="flex max-h-full w-full max-w-4xl flex-col rounded-2xl border border-slate-200 bg-white shadow-xl dark:border-slate-800 dark:bg-slate-900">
            <div className="flex items-center justify-between border-b border-slate-200 px-5 py-4 dark:border-slate-800">
              <div>
                <h2 className="text-lg font-semibold">
                  {compose.mode === 'new' ? '写信 / 发信' : compose.mode === 'reply' ? '回复' : compose.mode === 'replyAll' ? '回复全部' : '转发'}
                </h2>
                <div className="mt-1 text-xs text-slate-500">{activeMailbox?.email || '-'}</div>
              </div>
              <Button variant="ghost" size="icon" onClick={() => setCompose(EMPTY_COMPOSE)}>
                <X className="h-4 w-4" />
              </Button>
            </div>
            <div className="min-h-0 flex-1 space-y-4 overflow-y-auto px-5 py-4">
              <div className="grid gap-3">
                <Input value={compose.to} onChange={(event) => setCompose((current) => ({ ...current, to: event.target.value }))} placeholder="To" />
                <Input value={compose.cc} onChange={(event) => setCompose((current) => ({ ...current, cc: event.target.value }))} placeholder="Cc" />
                <Input value={compose.bcc} onChange={(event) => setCompose((current) => ({ ...current, bcc: event.target.value }))} placeholder="Bcc" />
                <Input value={compose.subject} onChange={(event) => setCompose((current) => ({ ...current, subject: event.target.value }))} placeholder={t('subject')} />
                <textarea
                  className="min-h-[16rem] w-full rounded-md border border-slate-200 bg-transparent px-3 py-3 text-sm shadow-sm dark:border-slate-800"
                  value={compose.bodyText}
                  onChange={(event) => setCompose((current) => ({ ...current, bodyText: event.target.value }))}
                />
                <div className="rounded-lg border border-dashed border-slate-300 px-3 py-3 dark:border-slate-700">
                  <div className="mb-2 text-sm font-medium">附件上传</div>
                  <input type="file" multiple onChange={(event) => void handleComposeFiles(event.target.files)} />
                  {compose.attachmentNames.length > 0 ? (
                    <div className="mt-3 flex flex-wrap gap-2">
                      {compose.attachmentNames.map((name, index) => (
                        <Badge key={`${name}-${index}`} variant="outline" className="gap-2">
                          {name}
                          <button
                            type="button"
                            onClick={() =>
                              setCompose((current) => ({
                                ...current,
                                attachmentNames: current.attachmentNames.filter((_, itemIndex) => itemIndex !== index),
                                attachments: current.attachments.filter((_, itemIndex) => itemIndex !== index),
                              }))
                            }
                          >
                            <X className="h-3 w-3" />
                          </button>
                        </Badge>
                      ))}
                    </div>
                  ) : null}
                </div>
              </div>
            </div>
            <div className="flex items-center justify-between border-t border-slate-200 px-5 py-4 dark:border-slate-800">
              <div className="text-xs text-slate-500">草稿保存支持继续编辑；回复、回复全部、转发也可以先保存草稿。</div>
              <div className="flex items-center gap-2">
                <Button variant="outline" onClick={() => void submitCompose(false)} disabled={compose.submitting}>
                  {compose.submitting ? <LoaderCircle className="mr-2 h-4 w-4 animate-spin" /> : null}
                  保存草稿
                </Button>
                <Button onClick={() => void submitCompose(true)} disabled={compose.submitting}>
                  {compose.submitting ? <LoaderCircle className="mr-2 h-4 w-4 animate-spin" /> : null}
                  发送
                </Button>
              </div>
            </div>
          </div>
        </div>
      ) : null}

      {searchOpen ? (
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/45 px-4 py-6">
          <div className="flex max-h-full w-full max-w-5xl flex-col rounded-2xl border border-slate-200 bg-white shadow-xl dark:border-slate-800 dark:bg-slate-900">
            <div className="flex items-center justify-between border-b border-slate-200 px-5 py-4 dark:border-slate-800">
              <div>
                <h2 className="text-lg font-semibold">跨邮箱统一搜索</h2>
                <div className="mt-1 text-xs text-slate-500">基于本地索引与缓存搜索，适合跨邮箱、标签、备注和会话定位。</div>
              </div>
              <Button variant="ghost" size="icon" onClick={() => setSearchOpen(false)}>
                <X className="h-4 w-4" />
              </Button>
            </div>
            <div className="space-y-4 px-5 py-4">
              <div className="flex flex-wrap items-center gap-3">
                <div className="min-w-[280px] flex-1">
                  <Input value={globalSearchQuery} onChange={(event) => setGlobalSearchQuery(event.target.value)} placeholder="输入主题、发件人、备注或标签关键词" />
                </div>
                <label className="flex items-center gap-2 text-sm">
                  <input type="checkbox" checked={!searchAcrossAll} onChange={(event) => setSearchAcrossAll(!event.target.checked)} />
                  仅当前邮箱
                </label>
                <Button onClick={() => void executeGlobalSearch(1)} disabled={globalSearchLoading}>
                  {globalSearchLoading ? <LoaderCircle className="mr-2 h-4 w-4 animate-spin" /> : <Search className="mr-2 h-4 w-4" />}
                  搜索
                </Button>
              </div>
              <div className="min-h-[22rem] overflow-y-auto rounded-xl border border-slate-200 dark:border-slate-800">
                {globalSearchResults.length === 0 ? (
                  <div className="flex h-full items-center justify-center px-6 py-12 text-sm text-slate-500">
                    {globalSearchLoading ? '正在搜索...' : '暂无搜索结果'}
                  </div>
                ) : (
                  <div className="divide-y divide-slate-200 dark:divide-slate-800">
                    {globalSearchResults.map((item) => (
                      <button key={`${item.mailboxId}-${item.id}`} onClick={() => jumpToMessage(item)} className="w-full px-4 py-3 text-left hover:bg-slate-50 dark:hover:bg-slate-900/60">
                        <div className="flex items-center justify-between gap-3">
                          <div className="min-w-0">
                            <div className="truncate text-sm font-medium">{item.subject}</div>
                            <div className="mt-1 truncate text-xs text-slate-500">
                              {item.mailboxEmail || '-'} · {item.sender} · {item.folderId}
                            </div>
                          </div>
                          <div className="text-[11px] text-slate-500">{item.date ? format(new Date(item.date), 'MM-dd HH:mm') : '-'}</div>
                        </div>
                        {item.meta.tags.length > 0 ? <div className="mt-2 text-xs text-slate-500">标签: {item.meta.tags.join(', ')}</div> : null}
                      </button>
                    ))}
                  </div>
                )}
              </div>
              <div className="flex items-center justify-end gap-2">
                <Button variant="ghost" size="sm" disabled={!globalSearchMeta.has_prev} onClick={() => void executeGlobalSearch(Math.max(1, globalSearchMeta.page - 1))}>
                  上一页
                </Button>
                <span className="text-xs text-slate-500">
                  {globalSearchMeta.page || 1} / {globalSearchMeta.total_pages || 1}
                </span>
                <Button variant="ghost" size="sm" disabled={!globalSearchMeta.has_next} onClick={() => void executeGlobalSearch(globalSearchMeta.page + 1)}>
                  下一页
                </Button>
              </div>
            </div>
          </div>
        </div>
      ) : null}

      {rulesOpen ? (
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/45 px-4 py-6">
          <div className="flex max-h-full w-full max-w-6xl flex-col rounded-2xl border border-slate-200 bg-white shadow-xl dark:border-slate-800 dark:bg-slate-900">
            <div className="flex items-center justify-between border-b border-slate-200 px-5 py-4 dark:border-slate-800">
              <div>
                <h2 className="text-lg font-semibold">规则引擎</h2>
                <div className="mt-1 text-xs text-slate-500">支持条件匹配、标签追加、标记已读、移动文件夹、跟进与备注追加。</div>
              </div>
              <Button variant="ghost" size="icon" onClick={() => setRulesOpen(false)}>
                <X className="h-4 w-4" />
              </Button>
            </div>
            <div className="grid min-h-0 flex-1 gap-4 overflow-hidden px-5 py-4 lg:grid-cols-[minmax(0,1fr)_minmax(360px,420px)]">
              <div className="min-h-0 overflow-y-auto rounded-xl border border-slate-200 dark:border-slate-800">
                <div className="flex items-center justify-between border-b border-slate-200 px-4 py-3 dark:border-slate-800">
                  <div className="flex items-center gap-2">
                    <ShieldCheck className="h-4 w-4" />
                    <span className="text-sm font-semibold">已保存规则</span>
                  </div>
                  <div className="flex gap-2">
                    <Button variant="outline" size="sm" onClick={openCreateRule}>
                      新建规则
                    </Button>
                    <Button variant="outline" size="sm" onClick={() => void handleApplyRules()} disabled={busyAction === 'rule-apply'}>
                      一键应用
                    </Button>
                  </div>
                </div>
                {loadingRules ? (
                  <div className="flex items-center justify-center gap-2 px-4 py-10 text-sm text-slate-500">
                    <LoaderCircle className="h-4 w-4 animate-spin" />
                    {t('loading')}
                  </div>
                ) : rules.length === 0 ? (
                  <div className="px-4 py-10 text-center text-sm text-slate-500">当前邮箱还没有规则</div>
                ) : (
                  <div className="space-y-3 p-4">
                    {rules.map((rule) => (
                      <div key={rule.id} className="rounded-xl border border-slate-200 p-4 dark:border-slate-800">
                        <div className="flex items-start justify-between gap-3">
                          <div>
                            <div className="text-sm font-semibold">{rule.name}</div>
                            <div className="mt-1 text-xs text-slate-500">priority: {rule.priority}</div>
                          </div>
                          <Badge variant={rule.enabled ? 'secondary' : 'outline'}>{rule.enabled ? 'enabled' : 'disabled'}</Badge>
                        </div>
                        <pre className="mt-3 whitespace-pre-wrap rounded-lg bg-slate-50 p-3 text-xs text-slate-600 dark:bg-slate-950 dark:text-slate-300">
                          {JSON.stringify(rule.conditions, null, 2)}
                        </pre>
                        <pre className="mt-3 whitespace-pre-wrap rounded-lg bg-slate-50 p-3 text-xs text-slate-600 dark:bg-slate-950 dark:text-slate-300">
                          {JSON.stringify(rule.actions, null, 2)}
                        </pre>
                        <div className="mt-3 flex flex-wrap gap-2">
                          <Button variant="ghost" size="sm" onClick={() => openEditRule(rule)}>
                            编辑
                          </Button>
                          <Button variant="ghost" size="sm" onClick={() => void handleApplyRules(rule.id)}>
                            应用当前规则
                          </Button>
                          <Button variant="destructive" size="sm" onClick={() => void handleDeleteRule(rule.id)}>
                            删除
                          </Button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div className="min-h-0 overflow-y-auto rounded-xl border border-slate-200 p-4 dark:border-slate-800">
                <h3 className="mb-4 text-sm font-semibold">{ruleEditor.id ? '编辑规则' : '新建规则'}</h3>
                <div className="grid gap-3">
                  <Input value={ruleEditor.name} onChange={(event) => setRuleEditor((current) => ({ ...current, name: event.target.value }))} placeholder="规则名称" />
                  <div className="grid grid-cols-[1fr_120px] gap-3">
                    <label className="flex items-center gap-2 rounded-md border border-slate-200 px-3 py-2 text-sm dark:border-slate-800">
                      <input
                        type="checkbox"
                        checked={ruleEditor.enabled}
                        onChange={(event) => setRuleEditor((current) => ({ ...current, enabled: event.target.checked }))}
                      />
                      启用规则
                    </label>
                    <Input
                      type="number"
                      value={ruleEditor.priority}
                      onChange={(event) => setRuleEditor((current) => ({ ...current, priority: Number(event.target.value) || 100 }))}
                    />
                  </div>
                  <div>
                    <label className="mb-1 block text-xs font-medium text-slate-500">条件 JSON</label>
                    <textarea
                      className="min-h-40 w-full rounded-md border border-slate-200 bg-transparent px-3 py-2 font-mono text-sm shadow-sm dark:border-slate-800"
                      value={ruleEditor.conditionsText}
                      onChange={(event) => setRuleEditor((current) => ({ ...current, conditionsText: event.target.value }))}
                    />
                  </div>
                  <div>
                    <label className="mb-1 block text-xs font-medium text-slate-500">动作 JSON</label>
                    <textarea
                      className="min-h-40 w-full rounded-md border border-slate-200 bg-transparent px-3 py-2 font-mono text-sm shadow-sm dark:border-slate-800"
                      value={ruleEditor.actionsText}
                      onChange={(event) => setRuleEditor((current) => ({ ...current, actionsText: event.target.value }))}
                    />
                  </div>
                  <div className="flex justify-end gap-2">
                    <Button variant="outline" onClick={() => setRuleEditor(EMPTY_RULE_EDITOR)}>
                      清空
                    </Button>
                    <Button onClick={() => void saveRuleEditor()} disabled={busyAction === 'rule-save'}>
                      {busyAction === 'rule-save' ? <LoaderCircle className="mr-2 h-4 w-4 animate-spin" /> : null}
                      保存规则
                    </Button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      ) : null}

      {opsOpen ? (
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/45 px-4 py-6">
          <div className="flex max-h-full w-full max-w-5xl flex-col rounded-2xl border border-slate-200 bg-white shadow-xl dark:border-slate-800 dark:bg-slate-900">
            <div className="flex items-center justify-between border-b border-slate-200 px-5 py-4 dark:border-slate-800">
              <div>
                <h2 className="text-lg font-semibold">审计日志和同步中心</h2>
                <div className="mt-1 text-xs text-slate-500">查看同步状态、任务历史和高频后台操作审计。</div>
              </div>
              <div className="flex items-center gap-2">
                <Button onClick={() => void handleRunSync()} disabled={busyAction === 'sync'}>
                  {busyAction === 'sync' ? <LoaderCircle className="mr-2 h-4 w-4 animate-spin" /> : <RefreshCw className="mr-2 h-4 w-4" />}
                  立即同步
                </Button>
                <Button variant="ghost" size="icon" onClick={() => setOpsOpen(false)}>
                  <X className="h-4 w-4" />
                </Button>
              </div>
            </div>
            <div className="grid min-h-0 flex-1 gap-4 overflow-hidden px-5 py-4 lg:grid-cols-2">
              <div className="min-h-0 overflow-y-auto rounded-xl border border-slate-200 p-4 dark:border-slate-800">
                <div className="mb-3 flex items-center justify-between">
                  <h3 className="text-sm font-semibold">同步中心</h3>
                  <Badge variant="outline">{syncStatus?.states.length || 0}</Badge>
                </div>
                {loadingOps ? (
                  <div className="flex items-center gap-2 text-sm text-slate-500">
                    <LoaderCircle className="h-4 w-4 animate-spin" />
                    {t('loading')}
                  </div>
                ) : !syncStatus ? (
                  <div className="text-sm text-slate-500">暂无同步数据</div>
                ) : (
                  <div className="space-y-3">
                    {syncStatus.jobs.map((job) => (
                      <div key={job.id} className="rounded-lg border border-slate-200 p-3 dark:border-slate-800">
                        <div className="flex items-center justify-between gap-3">
                          <span className="text-sm font-medium">Job #{job.id}</span>
                          <Badge variant={job.status === 'completed' ? 'secondary' : job.status === 'failed' ? 'destructive' : 'outline'}>
                            {job.status}
                          </Badge>
                        </div>
                        <div className="mt-2 text-xs text-slate-500">
                          <div>folders: {job.folders_synced}</div>
                          <div>cached: {job.cached_messages}</div>
                          <div>{job.started_at || '-'}</div>
                        </div>
                      </div>
                    ))}
                    {syncStatus.states.map((state) => (
                      <div key={`${state.method}-${state.folder_id}`} className="rounded-lg border border-slate-200 p-3 dark:border-slate-800">
                        <div className="flex items-center justify-between gap-3">
                          <span className="text-sm font-medium">{state.folder_id || 'INBOX'}</span>
                          <Badge variant={state.status === 'completed' ? 'secondary' : state.status === 'failed' ? 'destructive' : 'outline'}>
                            {state.status}
                          </Badge>
                        </div>
                        <div className="mt-2 text-xs text-slate-500">
                          <div>{methodLabel(state.method)}</div>
                          <div>cached: {state.cached_messages}</div>
                          <div>{state.last_synced_at || '-'}</div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <div className="min-h-0 overflow-y-auto rounded-xl border border-slate-200 p-4 dark:border-slate-800">
                <div className="mb-3 flex items-center justify-between">
                  <h3 className="text-sm font-semibold">审计日志</h3>
                  <Badge variant="outline">{auditLogs.length}</Badge>
                </div>
                {auditLogs.length === 0 ? (
                  <div className="text-sm text-slate-500">暂无审计日志</div>
                ) : (
                  <div className="space-y-3">
                    {auditLogs.map((item) => (
                      <div key={item.id} className="rounded-lg border border-slate-200 p-3 dark:border-slate-800">
                        <div className="flex items-center justify-between gap-3">
                          <div className="text-sm font-medium">{item.action}</div>
                          <Badge variant={item.status === 'success' ? 'secondary' : item.status === 'failed' ? 'destructive' : 'outline'}>
                            {item.status}
                          </Badge>
                        </div>
                        <div className="mt-2 text-xs text-slate-500">
                          <div>{item.target_type} · {item.target_id || '-'}</div>
                          <div>{item.created_at || '-'}</div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}
