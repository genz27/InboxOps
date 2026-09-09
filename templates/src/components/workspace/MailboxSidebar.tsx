import {
  ChevronRight,
  Copy,
  FolderPen,
  FolderPlus,
  LoaderCircle,
  Pin,
  PinOff,
  RefreshCw,
  Search,
  Settings2,
  ShieldCheck,
  Trash2,
  X,
} from 'lucide-react';
import type { MouseEvent } from 'react';
import { Badge } from '../ui/Badge';
import { Button } from '../ui/Button';
import { Input } from '../ui/Input';
import { cn } from '../../lib/utils';
import { methodLabel } from '../../lib/mail';
import type { EmailAccount, Folder } from '../../store/useAppStore';

export interface MailboxSidebarCapabilities {
  canManageFolders: boolean;
  writeUnsupportedHint: string;
}

interface MailboxSidebarProps {
  foldersCollapsed: boolean;
  otpNotifyEnabled: boolean;
  mailboxQuery: string;
  deferredMailboxQuery: string;
  accounts: EmailAccount[];
  filteredAccounts: EmailAccount[];
  pinnedIds: string[];
  activeMailboxId: string;
  loadingAccounts: boolean;
  loadingFolders: boolean;
  loadingLabel: string;
  foldersLabel: string;
  folders: Folder[];
  activeFolderId: string;
  activeFolder: Folder | null;
  capabilities: MailboxSidebarCapabilities;
  accountActionBusy: string;
  onToggleOtpNotify: () => void;
  onMailboxQueryChange: (value: string) => void;
  onSelectMailbox: (accountId: string) => void;
  onTogglePin: (accountId: string, event?: MouseEvent) => void;
  onCopyAccountEmail: (email: string, event?: MouseEvent) => void;
  onTestAccount: (accountId: string, event?: MouseEvent) => void;
  onOpenAccounts: (event: MouseEvent) => void;
  onToggleFoldersCollapsed: () => void;
  onCreateFolder: () => void;
  onRenameFolder: () => void;
  onDeleteFolder: () => void;
  onSelectFolder: (folderId: string) => void;
}

function MailboxList({
  loadingAccounts,
  loadingLabel,
  accounts,
  filteredAccounts,
  pinnedIds,
  activeMailboxId,
  accountActionBusy,
  onSelectMailbox,
  onTogglePin,
  onCopyAccountEmail,
  onTestAccount,
  onOpenAccounts,
}: Pick<
  MailboxSidebarProps,
  | 'loadingAccounts'
  | 'loadingLabel'
  | 'accounts'
  | 'filteredAccounts'
  | 'pinnedIds'
  | 'activeMailboxId'
  | 'accountActionBusy'
  | 'onSelectMailbox'
  | 'onTogglePin'
  | 'onCopyAccountEmail'
  | 'onTestAccount'
  | 'onOpenAccounts'
>) {
  if (loadingAccounts) {
    return (
      <div className="flex items-center gap-2 rounded-md px-2 py-2 text-sm text-white/45">
        <LoaderCircle className="h-4 w-4 animate-spin" />
        {loadingLabel}
      </div>
    );
  }
  if (accounts.length === 0) {
    return (
      <div className="rounded-md border border-dashed border-white/15 px-3 py-4 text-sm text-white/45">
        暂无邮箱档案
      </div>
    );
  }
  if (filteredAccounts.length === 0) {
    return (
      <div className="rounded-md border border-dashed border-white/15 px-3 py-4 text-sm text-white/45">
        无匹配邮箱
      </div>
    );
  }

  return (
    <>
      {filteredAccounts.map((account) => {
        const pinned = pinnedIds.includes(account.id);
        return (
          <div
            key={account.id}
            className={cn(
              'group rounded-lg border px-2 py-2 transition-colors',
              activeMailboxId === account.id
                ? 'border-white/20 bg-white/10'
                : 'border-transparent hover:bg-white/5',
            )}
          >
            <button type="button" onClick={() => onSelectMailbox(account.id)} className="w-full text-left">
              <div className="flex items-start justify-between gap-2">
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5">
                    {pinned ? <Pin className="h-3 w-3 shrink-0 text-amber-500" /> : null}
                    <div className="truncate text-sm font-medium">{account.label || account.email}</div>
                  </div>
                  <div className="truncate text-xs text-white/45">{account.email}</div>
                </div>
                <div className="flex shrink-0 items-center gap-1.5">
                  {(account.unreadCount ?? 0) > 0 ? (
                    <Badge variant="secondary" className="h-5 min-w-5 justify-center px-1 text-[10px]">
                      {account.unreadCount}
                    </Badge>
                  ) : null}
                  <div
                    className={cn(
                      'h-2.5 w-2.5 rounded-full',
                      account.status === 'connected'
                        ? 'bg-emerald-500'
                        : account.status === 'disconnected'
                          ? 'bg-red-500'
                          : 'bg-white/30',
                    )}
                  />
                </div>
              </div>
              <div className="mt-1.5 text-[11px] text-white/45">{methodLabel(account.preferredMethod)}</div>
            </button>
            <div className="mt-2 hidden flex-wrap gap-1 group-hover:flex">
              <Button
                variant="ghost"
                size="sm"
                className="h-7 px-2 text-[11px]"
                onClick={(event) => onTogglePin(account.id, event)}
                title={pinned ? '取消置顶' : '置顶'}
              >
                {pinned ? <PinOff className="h-3.5 w-3.5" /> : <Pin className="h-3.5 w-3.5" />}
              </Button>
              <Button
                variant="ghost"
                size="sm"
                className="h-7 px-2 text-[11px]"
                onClick={(event) => void onCopyAccountEmail(account.email, event)}
                title="复制邮箱"
              >
                <Copy className="h-3.5 w-3.5" />
              </Button>
              <Button
                variant="ghost"
                size="sm"
                className="h-7 px-2 text-[11px]"
                disabled={accountActionBusy === `test-${account.id}`}
                onClick={(event) => void onTestAccount(account.id, event)}
                title="测试连接"
              >
                {accountActionBusy === `test-${account.id}` ? (
                  <LoaderCircle className="h-3.5 w-3.5 animate-spin" />
                ) : (
                  <RefreshCw className="h-3.5 w-3.5" />
                )}
              </Button>
              <Button
                variant="ghost"
                size="sm"
                className="h-7 px-2 text-[11px]"
                onClick={onOpenAccounts}
                title="账号管理"
              >
                <Settings2 className="h-3.5 w-3.5" />
              </Button>
            </div>
          </div>
        );
      })}
    </>
  );
}

export function MailboxSidebar({
  foldersCollapsed,
  otpNotifyEnabled,
  mailboxQuery,
  deferredMailboxQuery,
  accounts,
  filteredAccounts,
  pinnedIds,
  activeMailboxId,
  loadingAccounts,
  loadingFolders,
  loadingLabel,
  foldersLabel,
  folders,
  activeFolderId,
  activeFolder,
  capabilities,
  accountActionBusy,
  onToggleOtpNotify,
  onMailboxQueryChange,
  onSelectMailbox,
  onTogglePin,
  onCopyAccountEmail,
  onTestAccount,
  onOpenAccounts,
  onToggleFoldersCollapsed,
  onCreateFolder,
  onRenameFolder,
  onDeleteFolder,
  onSelectFolder,
}: MailboxSidebarProps) {
  return (
    <>
      <div
        className={cn(
          'flex flex-col overflow-hidden p-4',
          foldersCollapsed
            ? 'min-h-0 flex-1'
            : 'max-h-[45%] min-h-[11rem] shrink-0 border-b border-white/10',
        )}
      >
        <div className="mb-2 flex items-center justify-between">
          <h2 className="text-sm font-semibold uppercase tracking-[0.24em] text-white/45">邮箱档案</h2>
          <div className="flex items-center gap-1">
            <Button
              variant="ghost"
              size="icon"
              className="h-7 w-7"
              title={otpNotifyEnabled ? '关闭验证码通知' : '开启验证码通知'}
              onClick={onToggleOtpNotify}
            >
              <ShieldCheck className={cn('h-3.5 w-3.5', otpNotifyEnabled ? 'text-emerald-500' : 'text-white/40')} />
            </Button>
            <Badge variant="outline">
              {filteredAccounts.length}
              {deferredMailboxQuery.trim() ? `/${accounts.length}` : ''}
            </Badge>
          </div>
        </div>
        <div className="relative mb-2">
          <Search className="absolute left-2.5 top-2 h-3.5 w-3.5 text-white/40" />
          <Input
            value={mailboxQuery}
            onChange={(event) => onMailboxQueryChange(event.target.value)}
            placeholder="搜索备注 / 邮箱"
            className="h-8 pl-8 text-xs"
          />
          {mailboxQuery ? (
            <button
              type="button"
              className="absolute right-2 top-1.5 rounded p-0.5 text-white/40 hover:text-white"
              onClick={() => onMailboxQueryChange('')}
              aria-label="清除搜索"
            >
              <X className="h-3.5 w-3.5" />
            </button>
          ) : null}
        </div>
        <div className="min-h-0 flex-1 space-y-1 overflow-y-auto pr-1">
          <MailboxList
            loadingAccounts={loadingAccounts}
            loadingLabel={loadingLabel}
            accounts={accounts}
            filteredAccounts={filteredAccounts}
            pinnedIds={pinnedIds}
            activeMailboxId={activeMailboxId}
            accountActionBusy={accountActionBusy}
            onSelectMailbox={onSelectMailbox}
            onTogglePin={onTogglePin}
            onCopyAccountEmail={onCopyAccountEmail}
            onTestAccount={onTestAccount}
            onOpenAccounts={onOpenAccounts}
          />
        </div>
      </div>

      <div
        className={cn(
          'flex min-h-0 flex-col overflow-hidden p-4',
          foldersCollapsed ? 'shrink-0 border-t border-white/10' : 'flex-1',
        )}
      >
        <div className={cn('flex items-center justify-between gap-2', foldersCollapsed ? 'mb-0' : 'mb-3')}>
          <button
            type="button"
            className="flex min-w-0 items-center gap-2 rounded-md text-left transition-colors hover:text-white"
            onClick={onToggleFoldersCollapsed}
            aria-expanded={!foldersCollapsed}
            aria-controls="workspace-folders-panel"
          >
            <ChevronRight
              className={cn(
                'h-4 w-4 shrink-0 text-white/40 transition-transform duration-200',
                !foldersCollapsed && 'rotate-90',
              )}
            />
            <h2 className="text-sm font-semibold uppercase tracking-[0.24em] text-white/45">{foldersLabel}</h2>
          </button>
          <div className="flex items-center gap-1">
            <Button
              variant="ghost"
              size="icon"
              className="h-7 w-7"
              disabled={!capabilities.canManageFolders}
              title={capabilities.canManageFolders ? '新建文件夹' : capabilities.writeUnsupportedHint}
              onClick={onCreateFolder}
            >
              <FolderPlus className="h-4 w-4" />
            </Button>
            <Button
              variant="ghost"
              size="icon"
              className="h-7 w-7"
              disabled={!capabilities.canManageFolders || !activeFolder || activeFolder.type !== 'custom'}
              title={capabilities.canManageFolders ? '重命名文件夹' : capabilities.writeUnsupportedHint}
              onClick={onRenameFolder}
            >
              <FolderPen className="h-4 w-4" />
            </Button>
            <Button
              variant="ghost"
              size="icon"
              className="h-7 w-7"
              disabled={!capabilities.canManageFolders || !activeFolder || activeFolder.type !== 'custom'}
              title={capabilities.canManageFolders ? '删除文件夹' : capabilities.writeUnsupportedHint}
              onClick={onDeleteFolder}
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
                <div className="flex items-center gap-2 rounded-md px-2 py-2 text-sm text-white/45">
                  <LoaderCircle className="h-4 w-4 animate-spin" />
                  {loadingLabel}
                </div>
              ) : (
                folders.map((folder) => (
                  <button
                    key={folder.id}
                    onClick={() => onSelectFolder(folder.id)}
                    className={cn(
                      'flex w-full items-center justify-between rounded-md px-2 py-2 text-left text-sm transition-colors',
                      activeFolderId === folder.id
                        ? 'bg-white/10 font-medium text-white'
                        : 'text-white/70 hover:bg-white/5',
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
    </>
  );
}
