import { create } from 'zustand';

export type CanonicalMethodValue = 'graph_api' | 'imap_new' | 'imap_old';
export type MethodValue = CanonicalMethodValue | 'Graph' | 'IMAP';
export type FolderType = 'inbox' | 'sent' | 'drafts' | 'trash' | 'spam' | 'archive' | 'custom';
export type AccountStatus = 'unknown' | 'connected' | 'disconnected';
export type EmailImportance = 'low' | 'normal' | 'high';
export type EmailStatus = 'active' | 'archived' | 'deleted' | 'snoozed' | 'done';

export interface MessageMeta {
  tags: string[];
  follow_up: string;
  notes: string;
  snoozed_until: string;
  status: EmailStatus;
  updated_at?: string;
}

export interface EmailAccount {
  id: string;
  label: string;
  email: string;
  method: MethodValue;
  preferredMethod?: MethodValue;
  notes: string;
  remark?: string;
  proxy?: string | null;
  clientId?: string;
  refreshToken?: string;
  status: AccountStatus;
  unreadCount?: number;
  totalCount?: number;
  lastSyncAt?: string;
  lastError?: string;
}

export interface Folder {
  id: string;
  name: string;
  displayName?: string;
  unreadCount: number;
  totalCount: number;
  type: FolderType;
  path?: string;
  parentId?: string | null;
  isSystem?: boolean;
}

export interface EmailAttachment {
  id: string;
  name: string;
  contentType?: string;
  type?: string;
  size: number;
  isInline?: boolean;
  is_inline?: boolean;
  contentId?: string;
}

export interface Email {
  id: string;
  mailboxId?: string;
  mailboxLabel?: string;
  mailboxEmail?: string;
  subject: string;
  sender: string;
  sender_name?: string;
  to_recipients: string[];
  cc_recipients: string[];
  bcc_recipients: string[];
  date: string;
  is_read: boolean;
  is_flagged: boolean;
  importance: EmailImportance;
  has_attachments: boolean;
  preview: string;
  body_text: string;
  body_html: string;
  headers: string;
  headersMap?: Record<string, string>;
  internet_message_id: string;
  conversation_id: string;
  attachments: EmailAttachment[];
  folderId: string;
  method?: MethodValue;
  meta?: MessageMeta;
  in_reply_to?: string;
  references?: string[];
}

interface AuthState {
  isAdmin: boolean;
  authReady: boolean;
  username: string | null;
  setAuthState: (authenticated: boolean, username?: string | null) => void;
  login: (username?: string | null) => void;
  logout: () => void;
}

interface WorkspaceState {
  accounts: EmailAccount[];
  activeMailboxId: string | null;
  folders: Folder[];
  emails: Email[];
  activeFolderId: string | null;
  activeEmailId: string | null;
  selectedEmailIds: string[];
  setAccounts: (accounts: EmailAccount[]) => void;
  setActiveMailboxId: (mailboxId: string | null) => void;
  setFolders: (folders: Folder[]) => void;
  setEmails: (emails: Email[]) => void;
  setActiveFolderId: (folderId: string | null) => void;
  setActiveEmailId: (emailId: string | null) => void;
  selectAllEmails: (emailIds: string[]) => void;
  toggleEmailSelection: (emailId: string) => void;
  clearEmailSelection: () => void;
  markAsRead: (emailIds: string[]) => void;
  markAsUnread: (emailIds: string[]) => void;
  starEmails: (emailIds: string[]) => void;
  unstarEmails: (emailIds: string[]) => void;
  deleteEmails: (emailIds: string[]) => void;
  moveEmails: (emailIds: string[], destinationFolderId: string) => void;
}

type AppState = AuthState & WorkspaceState;

function dedupeIds(ids: string[]): string[] {
  return Array.from(new Set(ids.filter((item) => Boolean(item))));
}

function filterSelection(ids: string[], emails: Email[]): string[] {
  const validIds = new Set(emails.map((email) => email.id));
  return ids.filter((id) => validIds.has(id));
}

function patchEmails(emails: Email[], targetIds: Set<string>, patch: Partial<Email>): Email[] {
  return emails.map((email) => (targetIds.has(email.id) ? { ...email, ...patch } : email));
}

export const useAppStore = create<AppState>((set) => ({
  isAdmin: false,
  authReady: false,
  username: null,
  accounts: [],
  activeMailboxId: null,
  folders: [],
  emails: [],
  activeFolderId: null,
  activeEmailId: null,
  selectedEmailIds: [],
  setAuthState: (authenticated, username = null) =>
    set({
      isAdmin: authenticated,
      authReady: true,
      username: authenticated ? username : null,
    }),
  login: (username = null) =>
    set({
      isAdmin: true,
      authReady: true,
      username,
    }),
  logout: () =>
    set({
      isAdmin: false,
      authReady: true,
      username: null,
    }),
  setAccounts: (accounts) =>
    set({
      accounts,
    }),
  setActiveMailboxId: (activeMailboxId) =>
    set({
      activeMailboxId,
    }),
  setFolders: (folders) =>
    set((state) => {
      const hasActiveFolder = state.activeFolderId ? folders.some((folder) => folder.id === state.activeFolderId) : false;
      return {
        folders,
        activeFolderId: hasActiveFolder ? state.activeFolderId : folders[0]?.id ?? null,
      };
    }),
  setEmails: (emails) =>
    set((state) => {
      const selectedEmailIds = filterSelection(state.selectedEmailIds, emails);
      const activeEmailStillExists = state.activeEmailId ? emails.some((email) => email.id === state.activeEmailId) : false;
      return {
        emails,
        selectedEmailIds,
        activeEmailId: activeEmailStillExists ? state.activeEmailId : null,
      };
    }),
  setActiveFolderId: (folderId) =>
    set({
      activeFolderId: folderId,
    }),
  setActiveEmailId: (emailId) =>
    set({
      activeEmailId: emailId,
    }),
  selectAllEmails: (emailIds) =>
    set({
      selectedEmailIds: dedupeIds(emailIds),
    }),
  toggleEmailSelection: (emailId) =>
    set((state) => {
      const selected = new Set(state.selectedEmailIds);
      if (selected.has(emailId)) {
        selected.delete(emailId);
      } else {
        selected.add(emailId);
      }
      return {
        selectedEmailIds: Array.from(selected),
      };
    }),
  clearEmailSelection: () =>
    set({
      selectedEmailIds: [],
    }),
  markAsRead: (emailIds) =>
    set((state) => ({
      emails: patchEmails(state.emails, new Set(emailIds), { is_read: true }),
    })),
  markAsUnread: (emailIds) =>
    set((state) => ({
      emails: patchEmails(state.emails, new Set(emailIds), { is_read: false }),
    })),
  starEmails: (emailIds) =>
    set((state) => ({
      emails: patchEmails(state.emails, new Set(emailIds), { is_flagged: true }),
    })),
  unstarEmails: (emailIds) =>
    set((state) => ({
      emails: patchEmails(state.emails, new Set(emailIds), { is_flagged: false }),
    })),
  deleteEmails: (emailIds) =>
    set((state) => {
      const targetIds = new Set(emailIds);
      const emails = state.emails.filter((email) => !targetIds.has(email.id));
      const selectedEmailIds = state.selectedEmailIds.filter((id) => !targetIds.has(id));
      return {
        emails,
        selectedEmailIds,
        activeEmailId: state.activeEmailId && targetIds.has(state.activeEmailId) ? null : state.activeEmailId,
      };
    }),
  moveEmails: (emailIds, destinationFolderId) =>
    set((state) => ({
      emails: patchEmails(state.emails, new Set(emailIds), { folderId: destinationFolderId }),
    })),
}));
