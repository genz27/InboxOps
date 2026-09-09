import React, { useDeferredValue, useEffect, useMemo, useRef, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { X } from 'lucide-react';
import { useI18n } from '../i18n';
import { Button } from '../components/ui/Button';
import { useFeedback } from '../components/Feedback';
import {
  type AuditLogRecord,
  type PaginationMeta,
  type RuleRecord,
  type SyncStatusRecord,
  applyRules,
  batchMessageAction,
  batchTestConnections,
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
import {
  accountMatchesQuery,
  copyText,
  emptyMeta,
  extractVerificationCodes,
  fileToAttachmentPayload,
  fromDatetimeLocalValue,
  methodCapabilities,
  toDatetimeLocalValue,
} from '../lib/mail';
import {
  ensureNotificationPermission,
  getPinnedMailboxIds,
  getRecentMailboxIds,
  isOtpNotifyEnabled,
  notifyVerificationCode,
  setOtpNotifyEnabled,
  sortMailboxesByPreference,
  togglePinnedMailbox,
  touchRecentMailbox,
} from '../lib/preferences';
import { isEditableTarget, resolveShortcut } from '../lib/shortcuts';
import { ComposeDrawer } from '../components/workspace/ComposeDrawer';
import { GlobalSearchModal } from '../components/workspace/GlobalSearchModal';
import { MailboxSidebar } from '../components/workspace/MailboxSidebar';
import { MailListPane } from '../components/workspace/MailListPane';
import { MessageReader } from '../components/workspace/MessageReader';
import { OpsModal } from '../components/workspace/OpsModal';
import { RulesModal } from '../components/workspace/RulesModal';
import { ShortcutsModal } from '../components/workspace/ShortcutsModal';
import {
  EMPTY_COMPOSE,
  EMPTY_META_FORM,
  EMPTY_RULE_EDITOR,
  formatMailDate,
  type ComposeFormState,
  type ComposeMode,
  type MetaFormState,
  type RuleEditorState,
  type ViewMode,
} from '../components/workspace/types';
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

type PendingFocus = {
  mailboxId: string;
  folderId: string;
  messageId: string;
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
  const receivedAt = formatMailDate(message.date);
  const quoted = message.body_text || message.preview || '';
  return `\n\n\n--- 原始邮件 ---\n主题: ${message.subject}\n发件人: ${message.sender}\n时间: ${receivedAt === '-' ? '' : receivedAt}\n\n${quoted}`;
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

export default function Workspace() {
  const { t } = useI18n();
  const navigate = useNavigate();
  const { toast, confirm, prompt } = useFeedback();
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
  const [toolsMenuOpen, setToolsMenuOpen] = useState(false);
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
  const [mailboxQuery, setMailboxQuery] = useState('');
  const deferredMailboxQuery = useDeferredValue(mailboxQuery);
  const [copiedCode, setCopiedCode] = useState('');
  const [pinnedIds, setPinnedIds] = useState<string[]>(() => getPinnedMailboxIds());
  const [recentIds, setRecentIds] = useState<string[]>(() => getRecentMailboxIds());
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [shortcutsOpen, setShortcutsOpen] = useState(false);
  const [otpNotifyEnabled, setOtpNotifyEnabledState] = useState(() => isOtpNotifyEnabled());
  const [accountActionBusy, setAccountActionBusy] = useState('');
  const listSearchRef = useRef<HTMLInputElement | null>(null);
  const toolsMenuRef = useRef<HTMLDivElement | null>(null);
  const seenMessageIdsRef = useRef<Set<string>>(new Set());
  const messagesRequestIdRef = useRef(0);
  const detailRequestIdRef = useRef(0);
  const foldersRequestIdRef = useRef(0);
  const loadingMessagesRef = useRef(false);
  const loadingDetailRef = useRef(false);
  const busyActionRef = useRef('');

  const activeMailbox = useMemo(
    () => accounts.find((account) => account.id === activeMailboxId) ?? null,
    [accounts, activeMailboxId],
  );
  const activeMethod: MethodValue = activeMailbox?.preferredMethod ?? 'graph_api';
  const capabilities = useMemo(() => methodCapabilities(activeMethod), [activeMethod]);
  const activeFolder = useMemo(() => folders.find((folder) => folder.id === activeFolderId) ?? null, [folders, activeFolderId]);
  const allVisibleSelected = messages.length > 0 && messages.every((message) => selectedMessageIds.includes(message.id));
  const syncStatus = syncItems[0] ?? null;
  const latestJob = syncStatus?.jobs?.[0] ?? null;
  const filteredAccounts = useMemo(() => {
    const matched = accounts.filter((account) => accountMatchesQuery(account, deferredMailboxQuery));
    return sortMailboxesByPreference(matched, pinnedIds, recentIds);
  }, [accounts, deferredMailboxQuery, pinnedIds, recentIds]);
  const verificationCodes = useMemo(
    () =>
      extractVerificationCodes(
        activeMessage?.subject,
        activeMessage?.preview,
        activeMessage?.body_text,
        activeMessage?.body_html ? activeMessage.body_html.replace(/<[^>]+>/g, ' ') : '',
      ),
    [activeMessage],
  );

  useEffect(() => {
    loadingMessagesRef.current = loadingMessages;
  }, [loadingMessages]);

  useEffect(() => {
    loadingDetailRef.current = loadingDetail;
  }, [loadingDetail]);

  useEffect(() => {
    busyActionRef.current = busyAction;
  }, [busyAction]);

  const syncFolders = useAppStore((state) => state.setFolders);
  const syncEmails = useAppStore((state) => state.setEmails);
  const syncActiveFolderId = useAppStore((state) => state.setActiveFolderId);
  const syncActiveEmailId = useAppStore((state) => state.setActiveEmailId);

  useEffect(() => {
    syncAccounts(accounts);
  }, [accounts, syncAccounts]);

  useEffect(() => {
    syncActiveMailboxId(activeMailboxId || null);
  }, [activeMailboxId, syncActiveMailboxId]);

  useEffect(() => {
    syncFolders(folders);
  }, [folders, syncFolders]);

  useEffect(() => {
    syncEmails(messages);
  }, [messages, syncEmails]);

  useEffect(() => {
    syncActiveFolderId(activeFolderId || null);
  }, [activeFolderId, syncActiveFolderId]);

  useEffect(() => {
    syncActiveEmailId(activeMessageId || null);
  }, [activeMessageId, syncActiveEmailId]);

  useEffect(() => {
    if (!toolsMenuOpen) {
      return;
    }
    const onPointerDown = (event: MouseEvent) => {
      if (!toolsMenuRef.current?.contains(event.target as Node)) {
        setToolsMenuOpen(false);
      }
    };
    window.addEventListener('pointerdown', onPointerDown);
    return () => window.removeEventListener('pointerdown', onPointerDown);
  }, [toolsMenuOpen]);

  async function loadAccounts() {
    setLoadingAccounts(true);
    try {
      const items = await listMailboxes();
      setAccounts(items);
      setActiveMailboxId((current) => {
        if (items.some((item) => item.id === current)) {
          return current;
        }
        const firstId = items[0]?.id ?? '';
        if (firstId) {
          setRecentIds(touchRecentMailbox(firstId));
        }
        return firstId;
      });
      if (items.length === 0) {
        setFolders([]);
        setMessages([]);
        setActiveFolderId('');
        setActiveMessageId('');
        setActiveMessage(null);
      }
    } catch (requestError) {
      updateError(requestError, '加载邮箱档案失败');
    } finally {
      setLoadingAccounts(false);
    }
  }

  async function loadFoldersForMailbox(mailboxId: string, method: MethodValue) {
    const requestId = foldersRequestIdRef.current + 1;
    foldersRequestIdRef.current = requestId;
    setLoadingFolders(true);
    try {
      const items = await listFolders(mailboxId, method);
      if (foldersRequestIdRef.current !== requestId) {
        return;
      }
      setFolders(items);
      setActiveFolderId((current) => {
        if (items.some((item) => item.id === current)) {
          return current;
        }
        return items.find((item) => item.type === 'inbox')?.id ?? items[0]?.id ?? 'INBOX';
      });
      const inbox = items.find((item) => item.type === 'inbox') ?? items[0];
      if (inbox) {
        setAccounts((current) =>
          current.map((account) =>
            account.id === mailboxId
              ? {
                  ...account,
                  unreadCount: inbox.unreadCount,
                  totalCount: inbox.totalCount,
                  lastSyncAt: new Date().toISOString(),
                }
              : account,
          ),
        );
      }
    } catch (requestError) {
      if (foldersRequestIdRef.current !== requestId) {
        return;
      }
      updateError(requestError, '加载文件夹失败');
      setFolders([]);
      setActiveFolderId('');
    } finally {
      if (foldersRequestIdRef.current === requestId) {
        setLoadingFolders(false);
      }
    }
  }

  async function loadMessagesForFolder(nextPage = page, options?: { silent?: boolean }) {
    if (!activeMailbox || !activeFolderId) {
      return;
    }
    const requestId = messagesRequestIdRef.current + 1;
    messagesRequestIdRef.current = requestId;
    const requestMailboxId = activeMailbox.id;
    const requestFolderId = activeFolderId;
    const requestMethod = activeMethod;
    const silent = options?.silent === true;
    if (!silent) {
      setLoadingMessages(true);
    }
    try {
      const response = await listMessages({
        mailboxId: requestMailboxId,
        method: requestMethod,
        folder: requestFolderId,
        page: nextPage,
        pageSize: ITEMS_PER_PAGE,
        keyword: deferredSearchQuery,
        unreadOnly: filterUnread,
        flaggedOnly: filterStarred,
        hasAttachmentsOnly: filterAttachment,
        sortOrder,
      });
      if (messagesRequestIdRef.current !== requestId) {
        return;
      }
      // 静默刷新时识别新验证码邮件并通知
      if (silent) {
        for (const item of response.items) {
          if (seenMessageIdsRef.current.has(item.id)) {
            continue;
          }
          seenMessageIdsRef.current.add(item.id);
          const codes = extractVerificationCodes(item.subject, item.preview, item.body_text);
          if (codes[0]) {
            notifyVerificationCode({
              title: `验证码 · ${item.subject || '新邮件'}`,
              body: `检测到验证码 ${codes[0]}（${item.sender || '未知发件人'}）`,
              code: codes[0],
            });
          }
        }
      } else {
        for (const item of response.items) {
          seenMessageIdsRef.current.add(item.id);
        }
      }

      setMessages(response.items);
      setListMeta(response.meta);
      setSelectedMessageIds((current) => current.filter((item) => response.items.some((message) => message.id === item)));
      if (!response.items.some((message) => message.id === activeMessageId) && !pendingFocus) {
        setActiveMessageId('');
        setActiveMessage(null);
        setThreadItems([]);
      }
      // 预取下一页，翻页更顺滑
      if (response.meta.has_next && !silent) {
        const prefetchPage = response.meta.page + 1;
        void listMessages({
          mailboxId: requestMailboxId,
          method: requestMethod,
          folder: requestFolderId,
          page: prefetchPage,
          pageSize: ITEMS_PER_PAGE,
          keyword: deferredSearchQuery,
          unreadOnly: filterUnread,
          flaggedOnly: filterStarred,
          hasAttachmentsOnly: filterAttachment,
          sortOrder,
        }).catch(() => {
          // 预取失败忽略
        });
      }
    } catch (requestError) {
      if (messagesRequestIdRef.current !== requestId) {
        return;
      }
      if (!silent) {
        updateError(requestError, '加载邮件失败');
        setMessages([]);
        setListMeta(EMPTY_PAGINATION);
      }
    } finally {
      if (!silent && messagesRequestIdRef.current === requestId) {
        setLoadingMessages(false);
      }
    }
  }

  async function loadMessageDetail(messageId: string, folderId = activeFolderId, preview?: Email | null) {
    if (!activeMailbox || !messageId) {
      return;
    }
    const requestId = detailRequestIdRef.current + 1;
    detailRequestIdRef.current = requestId;
    const requestMailboxId = activeMailbox.id;
    const requestMethod = activeMethod;
    setActiveMessageId(messageId);
    // 先展示列表摘要，减少空白等待
    if (preview && preview.id === messageId) {
      setActiveMessage(preview);
    }
    setLoadingDetail(true);
    try {
      const message = await getMessage({
        mailboxId: requestMailboxId,
        method: requestMethod,
        folder: folderId,
        messageId,
      });
      if (detailRequestIdRef.current !== requestId) {
        return;
      }
      setActiveMessage(message);
      // 打开未读邮件时自动标已读，减少一次手动点击
      if (!message.is_read && activeMailbox) {
        void updateReadState({
          mailboxId: requestMailboxId,
          method: requestMethod,
          folder: folderId,
          messageId,
          isRead: true,
        })
          .then((nextMessage) => {
            if (detailRequestIdRef.current !== requestId) {
              return;
            }
            updateMessageCollection(nextMessage);
          })
          .catch(() => {
            // 自动标已读失败不打断阅读
          });
      }
    } catch (requestError) {
      if (detailRequestIdRef.current !== requestId) {
        return;
      }
      updateError(requestError, '加载邮件详情失败');
    } finally {
      if (detailRequestIdRef.current === requestId) {
        setLoadingDetail(false);
      }
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
      updateError(requestError, '加载规则失败');
      setRules([]);
    } finally {
      setLoadingRules(false);
    }
  }

  useEffect(() => {
    void loadAccounts();
  }, []);

  // Esc / 邮件快捷键
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        if (compose.open) {
          setCompose(EMPTY_COMPOSE);
          return;
        }
        if (searchOpen) {
          setSearchOpen(false);
          return;
        }
        if (rulesOpen) {
          setRulesOpen(false);
          return;
        }
        if (opsOpen) {
          setOpsOpen(false);
          return;
        }
        if (toolsMenuOpen) {
          setToolsMenuOpen(false);
          return;
        }
        if (shortcutsOpen) {
          setShortcutsOpen(false);
          return;
        }
        if (mobileNavOpen) {
          setMobileNavOpen(false);
          return;
        }
        if (activeMessageId) {
          setActiveMessageId('');
          setActiveMessage(null);
        }
        return;
      }

      if (isEditableTarget(event.target) || compose.open || searchOpen || rulesOpen || opsOpen) {
        return;
      }

      const action = resolveShortcut(event);
      if (!action) {
        return;
      }
      event.preventDefault();

      if (action === 'focusSearch') {
        listSearchRef.current?.focus();
        return;
      }
      if (action === 'compose') {
        openCompose('new');
        return;
      }
      if (action === 'refresh') {
        void refreshEverything();
        return;
      }
      if (action === 'toggleUnreadFilter') {
        setFilterUnread((current) => !current);
        setPage(1);
        return;
      }
      if (action === 'copyOtp') {
        const code = verificationCodes[0];
        if (code) {
          void handleCopyVerificationCode(code);
        }
        return;
      }
      if (action === 'archive' && activeMessage) {
        void handleBatchAction('archive', [activeMessage.id]);
        return;
      }
      if (action === 'delete' && activeMessage) {
        void handleDeleteActiveMessage();
        return;
      }
      if (action === 'nextMessage' || action === 'prevMessage') {
        if (messages.length === 0) {
          return;
        }
        const currentIndex = Math.max(
          0,
          messages.findIndex((item) => item.id === activeMessageId),
        );
        const nextIndex =
          action === 'nextMessage'
            ? Math.min(messages.length - 1, (activeMessageId ? currentIndex : -1) + 1)
            : Math.max(0, (activeMessageId ? currentIndex : 0) - 1);
        const target = messages[nextIndex];
        if (target) {
          void loadMessageDetail(target.id, target.folderId, target);
        }
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [
    compose.open,
    searchOpen,
    rulesOpen,
    opsOpen,
    toolsMenuOpen,
    shortcutsOpen,
    mobileNavOpen,
    activeMessageId,
    activeMessage,
    messages,
    verificationCodes,
  ]);

  useEffect(() => {
    if (otpNotifyEnabled) {
      void ensureNotificationPermission();
    }
  }, [otpNotifyEnabled]);

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
      if (
        document.visibilityState === 'hidden' ||
        loadingMessagesRef.current ||
        loadingDetailRef.current ||
        busyActionRef.current !== ''
      ) {
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
    if (page !== 1) {
      setPage(1);
      await Promise.all([
        loadFoldersForMailbox(activeMailbox.id, activeMethod),
        loadOpsForMailbox(activeMailbox.id),
      ]);
      return;
    }
    await Promise.all([
      loadFoldersForMailbox(activeMailbox.id, activeMethod),
      loadMessagesForFolder(1),
      loadOpsForMailbox(activeMailbox.id),
    ]);
  }

  async function handleCopyVerificationCode(code: string) {
    try {
      await copyText(code);
      setCopiedCode(code);
      updateNotice(`验证码已复制：${code}`);
      window.setTimeout(() => {
        setCopiedCode((current) => (current === code ? '' : current));
      }, 1500);
    } catch {
      updateError(new Error('复制失败'), '复制验证码失败');
    }
  }

  function selectMailbox(accountId: string) {
    if (accountId === activeMailboxId) {
      setMobileNavOpen(false);
      return;
    }
    messagesRequestIdRef.current += 1;
    detailRequestIdRef.current += 1;
    foldersRequestIdRef.current += 1;
    seenMessageIdsRef.current = new Set();
    setActiveMailboxId(accountId);
    setRecentIds(touchRecentMailbox(accountId));
    setFolders([]);
    setActiveFolderId('');
    setMessages([]);
    setListMeta(EMPTY_PAGINATION);
    setPage(1);
    setSelectedMessageIds([]);
    setActiveMessageId('');
    setActiveMessage(null);
    setThreadItems([]);
    setMobileNavOpen(false);
  }

  function handleTogglePin(accountId: string, event?: React.MouseEvent) {
    event?.stopPropagation();
    setPinnedIds(togglePinnedMailbox(accountId));
  }

  async function handleCopyAccountEmail(email: string, event?: React.MouseEvent) {
    event?.stopPropagation();
    try {
      await copyText(email);
      updateNotice(`已复制邮箱：${email}`);
    } catch {
      updateError(new Error('复制失败'), '复制邮箱失败');
    }
  }

  async function handleTestAccount(accountId: string, event?: React.MouseEvent) {
    event?.stopPropagation();
    setAccountActionBusy(`test-${accountId}`);
    try {
      const payload = await batchTestConnections([accountId]);
      const first = payload.results[0];
      const ok = first?.success === true;
      setAccounts((current) =>
        current.map((account) =>
          account.id === accountId
            ? { ...account, status: ok ? 'connected' : 'disconnected', lastError: ok ? '' : String(first?.message || '') }
            : account,
        ),
      );
      updateNotice(ok ? '连接测试成功' : `连接失败：${String(first?.message || '未知错误')}`);
    } catch (requestError) {
      updateError(requestError, '连接测试失败');
    } finally {
      setAccountActionBusy('');
    }
  }

  function updateError(requestError: unknown, fallback: string) {
    toast(requestError instanceof Error ? requestError.message : fallback, 'error');
  }

  function updateNotice(message: string) {
    toast(message);
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
      const clearsActiveMessage = action === 'delete' || action === 'archive' || action === 'move';
      if (clearsActiveMessage && messageIds.includes(activeMessageId)) {
        setActiveMessage(null);
        setActiveMessageId('');
      }
      setSelectedMessageIds((current) => current.filter((item) => !messageIds.includes(item)));
      if (!clearsActiveMessage && activeMessage && messageIds.includes(activeMessage.id)) {
        const nextRead =
          action === 'mark_read' ? true : action === 'mark_unread' ? false : activeMessage.is_read;
        const nextFlagged =
          action === 'flag' ? true : action === 'unflag' ? false : activeMessage.is_flagged;
        updateMessageCollection({
          ...activeMessage,
          is_read: nextRead,
          is_flagged: nextFlagged,
        });
      }
      updateNotice(`操作完成，成功 ${payload.summary.succeeded} 条，失败 ${payload.summary.failed} 条`);
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
    if (!capabilities.canCompose) {
      updateError(new Error(capabilities.writeUnsupportedHint), capabilities.writeUnsupportedHint);
      return;
    }
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
    if (!capabilities.canSaveDraft) {
      updateError(new Error(capabilities.writeUnsupportedHint), capabilities.writeUnsupportedHint);
      return;
    }
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
    if (!capabilities.canManageFolders) {
      updateError(new Error(capabilities.writeUnsupportedHint), capabilities.writeUnsupportedHint);
      return;
    }
    try {
      if (type === 'create') {
        const name = await prompt({ title: '新建文件夹', label: '文件夹名称', confirmLabel: '创建' });
        if (!name) {
          return;
        }
        await createFolder({ mailboxId: activeMailbox.id, method: activeMethod, displayName: name });
        updateNotice('文件夹已创建');
      }
      if (type === 'rename' && activeFolder) {
        const name = await prompt({
          title: '重命名文件夹',
          label: '新名称',
          defaultValue: activeFolder.displayName,
          confirmLabel: '保存',
        });
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
        const ok = await confirm({
          title: '删除文件夹',
          body: `确认删除「${activeFolder.displayName}」？文件夹内邮件可能一并受影响。`,
          confirmLabel: '删除',
          danger: true,
        });
        if (!ok) {
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
    if (!activeMailbox) {
      return;
    }
    const ok = await confirm({
      title: '删除规则',
      body: '确认删除这条规则？此操作不可恢复。',
      confirmLabel: '删除',
      danger: true,
    });
    if (!ok) {
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

  const mailboxSidebar = (
    <MailboxSidebar
      foldersCollapsed={foldersCollapsed}
      otpNotifyEnabled={otpNotifyEnabled}
      mailboxQuery={mailboxQuery}
      deferredMailboxQuery={deferredMailboxQuery}
      accounts={accounts}
      filteredAccounts={filteredAccounts}
      pinnedIds={pinnedIds}
      activeMailboxId={activeMailboxId}
      loadingAccounts={loadingAccounts}
      loadingFolders={loadingFolders}
      loadingLabel={t('loading')}
      foldersLabel={t('folders')}
      folders={folders}
      activeFolderId={activeFolderId}
      activeFolder={activeFolder}
      capabilities={capabilities}
      accountActionBusy={accountActionBusy}
      onToggleOtpNotify={() => {
        const next = !otpNotifyEnabled;
        setOtpNotifyEnabledState(next);
        setOtpNotifyEnabled(next);
        if (next) {
          void ensureNotificationPermission();
        }
      }}
      onMailboxQueryChange={setMailboxQuery}
      onSelectMailbox={selectMailbox}
      onTogglePin={handleTogglePin}
      onCopyAccountEmail={handleCopyAccountEmail}
      onTestAccount={handleTestAccount}
      onOpenAccounts={(event) => {
        event.stopPropagation();
        navigate('/accounts');
      }}
      onToggleFoldersCollapsed={() => setFoldersCollapsed((current) => !current)}
      onCreateFolder={() => void handleFolderAction('create')}
      onRenameFolder={() => void handleFolderAction('rename')}
      onDeleteFolder={() => void handleFolderAction('delete')}
      onSelectFolder={(folderId) => {
        setActiveFolderId(folderId);
        setPage(1);
        setSelectedMessageIds([]);
        setMobileNavOpen(false);
      }}
    />
  );

  return (
    <div className="relative flex h-full w-full overflow-hidden bg-black">
      <div className="hidden w-72 shrink-0 overflow-hidden border-r border-white/10 bg-[#0a0a0a] lg:flex lg:flex-col">
        {mailboxSidebar}
      </div>

      {mobileNavOpen ? (
        <div className="absolute inset-0 z-40 flex lg:hidden">
          <button type="button" className="absolute inset-0 bg-black/40" onClick={() => setMobileNavOpen(false)} aria-label="关闭侧栏" />
          <div className="relative z-10 flex h-full w-[18rem] max-w-[85vw] flex-col overflow-hidden border-r border-white/10 bg-[#0a0a0a] shadow-xl">
            <div className="flex items-center justify-between border-b border-white/10 px-3 py-2">
              <span className="text-sm font-semibold">邮箱与文件夹</span>
              <Button variant="ghost" size="icon" className="h-8 w-8" onClick={() => setMobileNavOpen(false)}>
                <X className="h-4 w-4" />
              </Button>
            </div>
            <div className="flex min-h-0 flex-1 flex-col overflow-hidden">{mailboxSidebar}</div>
          </div>
        </div>
      ) : null}

      <MailListPane
        labels={{
          search: t('search'),
          refresh: t('refresh'),
          compose: t('compose'),
          unread: t('unread'),
          starred: t('starred'),
          hasAttachment: t('hasAttachment'),
          dateDesc: t('dateDesc'),
          dateAsc: t('dateAsc'),
          clearSelection: t('clearSelection'),
          markRead: t('markRead'),
          markUnread: t('markUnread'),
          star: t('star'),
          archive: t('archive'),
          move: t('move'),
          delete: t('delete'),
          loading: t('loading'),
          noEmails: t('noEmails'),
          selectAll: t('selectAll'),
          page: t('page'),
        }}
        searchQuery={searchQuery}
        listSearchRef={listSearchRef}
        toolsMenuRef={toolsMenuRef}
        loadingMessages={loadingMessages}
        busyAction={busyAction}
        capabilities={capabilities}
        hasActiveMailbox={Boolean(activeMailbox)}
        filterUnread={filterUnread}
        filterStarred={filterStarred}
        filterAttachment={filterAttachment}
        sortOrder={sortOrder}
        toolsMenuOpen={toolsMenuOpen}
        selectedMessageIds={selectedMessageIds}
        latestJob={latestJob}
        messages={messages}
        deferredSearchQuery={deferredSearchQuery}
        activeMessageId={activeMessageId}
        allVisibleSelected={allVisibleSelected}
        listMeta={listMeta}
        onOpenMobileNav={() => setMobileNavOpen(true)}
        onSearchQueryChange={(value) => {
          setSearchQuery(value);
          setPage(1);
        }}
        onRefresh={() => void refreshEverything()}
        onCompose={() => openCompose('new')}
        onOpenShortcuts={() => setShortcutsOpen(true)}
        onToggleUnread={() => {
          setFilterUnread((current) => !current);
          setPage(1);
        }}
        onToggleStarred={() => {
          setFilterStarred((current) => !current);
          setPage(1);
        }}
        onToggleAttachment={() => {
          setFilterAttachment((current) => !current);
          setPage(1);
        }}
        onToggleSortOrder={() => {
          setSortOrder((current) => (current === 'desc' ? 'asc' : 'desc'));
          setPage(1);
        }}
        onToggleToolsMenu={() => setToolsMenuOpen((open) => !open)}
        onOpenSearch={() => {
          setToolsMenuOpen(false);
          setSearchOpen(true);
        }}
        onOpenRules={() => {
          setToolsMenuOpen(false);
          setRulesOpen(true);
        }}
        onOpenOps={() => {
          setToolsMenuOpen(false);
          setOpsOpen(true);
        }}
        onClearSelection={() => setSelectedMessageIds([])}
        onMarkRead={() => void handleBatchAction('mark_read', selectedMessageIds)}
        onMarkUnread={() => void handleBatchAction('mark_unread', selectedMessageIds)}
        onStar={() => void handleBatchAction('flag', selectedMessageIds)}
        onArchive={() => void handleBatchAction('archive', selectedMessageIds)}
        onMove={() => {
          void (async () => {
            const destinationFolder = await prompt({
              title: '移动到文件夹',
              label: '文件夹 ID 或名称',
              confirmLabel: '移动',
            });
            if (destinationFolder) {
              await handleBatchAction('move', selectedMessageIds, destinationFolder);
            }
          })();
        }}
        onDelete={() => {
          void (async () => {
            const ok = await confirm({
              title: '删除邮件',
              body: `确认删除已选的 ${selectedMessageIds.length} 封邮件？`,
              confirmLabel: '删除',
              danger: true,
            });
            if (ok) {
              await handleBatchAction('delete', selectedMessageIds);
            }
          })();
        }}
        onClearFilters={() => {
          setSearchQuery('');
          setFilterUnread(false);
          setFilterStarred(false);
          setFilterAttachment(false);
          setPage(1);
        }}
        onOpenMessage={(message) => void loadMessageDetail(message.id, message.folderId, message)}
        onToggleMessageSelected={(messageId) =>
          setSelectedMessageIds((current) =>
            current.includes(messageId) ? current.filter((item) => item !== messageId) : [...current, messageId],
          )
        }
        onToggleSelectAll={(checked) => setSelectedMessageIds(checked ? messages.map((item) => item.id) : [])}
        onPrevPage={() => setPage((current) => Math.max(1, current - 1))}
        onNextPage={() => setPage((current) => current + 1)}
      />

      <MessageReader
        labels={{
          loading: t('loading'),
          markUnread: t('markUnread'),
          markRead: t('markRead'),
          unstar: t('unstar'),
          star: t('star'),
          archive: t('archive'),
          delete: t('delete'),
          to: t('to'),
          cc: t('cc'),
          htmlView: t('htmlView'),
          textView: t('textView'),
          headersView: t('headersView'),
          attachments: t('attachments'),
          messageId: t('messageId'),
          conversationId: t('conversationId'),
          selectEmail: t('selectEmail'),
        }}
        loadingDetail={loadingDetail}
        activeMessage={activeMessage}
        activeMessageId={activeMessageId}
        activeFolderId={activeFolderId}
        activeMailbox={activeMailbox}
        activeMethod={activeMethod}
        capabilities={capabilities}
        verificationCodes={verificationCodes}
        copiedCode={copiedCode}
        viewMode={viewMode}
        detailSidebarCollapsed={detailSidebarCollapsed}
        metaPanelCollapsed={metaPanelCollapsed}
        threadPanelCollapsed={threadPanelCollapsed}
        metaForm={metaForm}
        busyAction={busyAction}
        threadItems={threadItems}
        onCloseMessage={() => setActiveMessageId('')}
        onEditDraft={openDraftEditor}
        onReply={(message) => openCompose('reply', message)}
        onReplyAll={(message) => openCompose('replyAll', message)}
        onForward={(message) => openCompose('forward', message)}
        onToggleRead={(isRead) => void handleSingleStateUpdate('read', isRead)}
        onToggleFlag={(isFlagged) => void handleSingleStateUpdate('flag', isFlagged)}
        onArchive={(messageId) => void handleBatchAction('archive', [messageId])}
        onDelete={() => void handleDeleteActiveMessage()}
        onCopyVerificationCode={handleCopyVerificationCode}
        onViewModeChange={setViewMode}
        onToggleDetailSidebar={() => setDetailSidebarCollapsed((current) => !current)}
        onToggleMetaPanel={() => setMetaPanelCollapsed((current) => !current)}
        onToggleThreadPanel={() => setThreadPanelCollapsed((current) => !current)}
        onMetaFormChange={(patch) => setMetaForm((current) => ({ ...current, ...patch }))}
        onSaveMeta={() => void handleSaveMeta()}
        onOpenThreadMessage={(message) => void loadMessageDetail(message.id, message.folderId, message)}
        onDownloadAttachment={handleDownloadAttachment}
      />

      {compose.open ? (
        <ComposeDrawer
          compose={compose}
          mailboxEmail={activeMailbox?.email || '-'}
          subjectPlaceholder={t('subject')}
          onClose={() => setCompose(EMPTY_COMPOSE)}
          onChange={(patch) => setCompose((current) => ({ ...current, ...patch }))}
          onFiles={(files) => void handleComposeFiles(files)}
          onRemoveAttachment={(index) =>
            setCompose((current) => ({
              ...current,
              attachmentNames: current.attachmentNames.filter((_, itemIndex) => itemIndex !== index),
              attachments: current.attachments.filter((_, itemIndex) => itemIndex !== index),
            }))
          }
          onSaveDraft={() => void submitCompose(false)}
          onSend={() => void submitCompose(true)}
        />
      ) : null}

      {shortcutsOpen ? <ShortcutsModal onClose={() => setShortcutsOpen(false)} /> : null}

      {searchOpen ? (
        <GlobalSearchModal
          query={globalSearchQuery}
          searchAcrossAll={searchAcrossAll}
          loading={globalSearchLoading}
          results={globalSearchResults}
          meta={globalSearchMeta}
          onQueryChange={setGlobalSearchQuery}
          onSearchAcrossAllChange={setSearchAcrossAll}
          onSearch={(nextPage) => void executeGlobalSearch(nextPage ?? 1)}
          onJump={jumpToMessage}
          onClose={() => setSearchOpen(false)}
        />
      ) : null}

      {rulesOpen ? (
        <RulesModal
          rules={rules}
          loading={loadingRules}
          loadingLabel={t('loading')}
          busyAction={busyAction}
          editor={ruleEditor}
          onEditorChange={setRuleEditor}
          onClose={() => setRulesOpen(false)}
          onCreate={openCreateRule}
          onApplyAll={() => void handleApplyRules()}
          onEdit={openEditRule}
          onApply={(ruleId) => void handleApplyRules(ruleId)}
          onDelete={(ruleId) => void handleDeleteRule(ruleId)}
          onSave={() => void saveRuleEditor()}
        />
      ) : null}

      {opsOpen ? (
        <OpsModal
          loading={loadingOps}
          loadingLabel={t('loading')}
          busyAction={busyAction}
          syncStatus={syncStatus}
          auditLogs={auditLogs}
          onSync={() => void handleRunSync()}
          onClose={() => setOpsOpen(false)}
        />
      ) : null}
    </div>
  );
}
