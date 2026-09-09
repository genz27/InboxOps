import {
  AlertCircle,
  ChevronLeft,
  ChevronRight,
  Clock,
  FilePlus2,
  Keyboard,
  LoaderCircle,
  Mail,
  Menu,
  MoreHorizontal,
  Paperclip,
  RefreshCw,
  Search,
  Star,
  X,
} from 'lucide-react';
import type { RefObject } from 'react';
import type { PaginationMeta, SyncJobRecord } from '../../lib/api';
import type { Email } from '../../store/useAppStore';
import { Badge } from '../ui/Badge';
import { Button } from '../ui/Button';
import { Input } from '../ui/Input';
import { cn } from '../../lib/utils';
import { formatMailDate } from './types';

export interface MailListPaneCapabilities {
  canCompose: boolean;
  writeUnsupportedHint: string;
}

interface MailListPaneLabels {
  search: string;
  refresh: string;
  compose: string;
  unread: string;
  starred: string;
  hasAttachment: string;
  dateDesc: string;
  dateAsc: string;
  clearSelection: string;
  markRead: string;
  markUnread: string;
  star: string;
  archive: string;
  move: string;
  delete: string;
  loading: string;
  noEmails: string;
  selectAll: string;
  page: string;
}

interface MailListPaneProps {
  labels: MailListPaneLabels;
  searchQuery: string;
  listSearchRef: RefObject<HTMLInputElement | null>;
  toolsMenuRef: RefObject<HTMLDivElement | null>;
  loadingMessages: boolean;
  busyAction: string;
  capabilities: MailListPaneCapabilities;
  hasActiveMailbox: boolean;
  filterUnread: boolean;
  filterStarred: boolean;
  filterAttachment: boolean;
  sortOrder: 'asc' | 'desc';
  toolsMenuOpen: boolean;
  selectedMessageIds: string[];
  latestJob: SyncJobRecord | null;
  messages: Email[];
  deferredSearchQuery: string;
  activeMessageId: string;
  allVisibleSelected: boolean;
  listMeta: PaginationMeta;
  onOpenMobileNav: () => void;
  onSearchQueryChange: (value: string) => void;
  onRefresh: () => void;
  onCompose: () => void;
  onOpenShortcuts: () => void;
  onToggleUnread: () => void;
  onToggleStarred: () => void;
  onToggleAttachment: () => void;
  onToggleSortOrder: () => void;
  onToggleToolsMenu: () => void;
  onOpenSearch: () => void;
  onOpenRules: () => void;
  onOpenOps: () => void;
  onClearSelection: () => void;
  onMarkRead: () => void;
  onMarkUnread: () => void;
  onStar: () => void;
  onArchive: () => void;
  onMove: () => void;
  onDelete: () => void;
  onClearFilters: () => void;
  onOpenMessage: (message: Email) => void;
  onToggleMessageSelected: (messageId: string) => void;
  onToggleSelectAll: (checked: boolean) => void;
  onPrevPage: () => void;
  onNextPage: () => void;
}

export function MailListPane({
  labels,
  searchQuery,
  listSearchRef,
  toolsMenuRef,
  loadingMessages,
  busyAction,
  capabilities,
  hasActiveMailbox,
  filterUnread,
  filterStarred,
  filterAttachment,
  sortOrder,
  toolsMenuOpen,
  selectedMessageIds,
  latestJob,
  messages,
  deferredSearchQuery,
  activeMessageId,
  allVisibleSelected,
  listMeta,
  onOpenMobileNav,
  onSearchQueryChange,
  onRefresh,
  onCompose,
  onOpenShortcuts,
  onToggleUnread,
  onToggleStarred,
  onToggleAttachment,
  onToggleSortOrder,
  onToggleToolsMenu,
  onOpenSearch,
  onOpenRules,
  onOpenOps,
  onClearSelection,
  onMarkRead,
  onMarkUnread,
  onStar,
  onArchive,
  onMove,
  onDelete,
  onClearFilters,
  onOpenMessage,
  onToggleMessageSelected,
  onToggleSelectAll,
  onPrevPage,
  onNextPage,
}: MailListPaneProps) {
  return (
    <div
      className={cn(
        'flex w-full shrink-0 flex-col border-r border-white/10 bg-black md:w-[28rem]',
        activeMessageId ? 'hidden md:flex' : 'flex',
      )}
    >
      <div className="space-y-3 border-b border-white/10 p-3">
        <div className="flex flex-wrap items-center gap-2">
          <Button variant="outline" size="icon" className="h-9 w-9 lg:hidden" onClick={onOpenMobileNav} title="打开邮箱列表">
            <Menu className="h-4 w-4" />
          </Button>
          <div className="relative min-w-[180px] flex-1">
            <Search className="absolute left-2.5 top-2 h-4 w-4 text-white/40" />
            <Input
              ref={listSearchRef}
              value={searchQuery}
              onChange={(event) => onSearchQueryChange(event.target.value)}
              placeholder={`${labels.search}  (/)`}
              className="pl-8"
            />
          </div>
          <Button variant="outline" onClick={onRefresh} disabled={loadingMessages || busyAction !== ''}>
            <RefreshCw className={cn('mr-2 h-4 w-4', loadingMessages && 'animate-spin')} />
            {labels.refresh}
          </Button>
          <Button
            onClick={onCompose}
            disabled={!capabilities.canCompose || !hasActiveMailbox}
            title={capabilities.canCompose ? `${labels.compose} (n)` : capabilities.writeUnsupportedHint}
          >
            <FilePlus2 className="mr-2 h-4 w-4" />
            {labels.compose}
          </Button>
          <Button variant="ghost" size="icon" className="h-9 w-9" title="快捷键帮助" onClick={onOpenShortcuts}>
            <Keyboard className="h-4 w-4" />
          </Button>
        </div>

        <div className="flex flex-wrap items-center gap-1">
          <Button
            variant={filterUnread ? 'secondary' : 'ghost'}
            size="icon"
            className="h-8 w-8"
            title={`${labels.unread} (u)`}
            onClick={onToggleUnread}
          >
            <Mail className="h-4 w-4" />
          </Button>
          <Button
            variant={filterStarred ? 'secondary' : 'ghost'}
            size="icon"
            className="h-8 w-8"
            title={labels.starred}
            onClick={onToggleStarred}
          >
            <Star className={cn('h-4 w-4', filterStarred && 'fill-amber-400 text-amber-400')} />
          </Button>
          <Button
            variant={filterAttachment ? 'secondary' : 'ghost'}
            size="icon"
            className="h-8 w-8"
            title={labels.hasAttachment}
            onClick={onToggleAttachment}
          >
            <Paperclip className="h-4 w-4" />
          </Button>
          <Button
            variant="ghost"
            size="icon"
            className="h-8 w-8"
            title={sortOrder === 'desc' ? labels.dateDesc : labels.dateAsc}
            onClick={onToggleSortOrder}
          >
            <Clock className="h-4 w-4" />
          </Button>
          <div className="relative ml-auto" ref={toolsMenuRef}>
            <Button
              variant={toolsMenuOpen ? 'secondary' : 'ghost'}
              size="icon"
              className="h-8 w-8"
              title="更多工具"
              onClick={onToggleToolsMenu}
            >
              <MoreHorizontal className="h-4 w-4" />
            </Button>
            {toolsMenuOpen ? (
              <div className="absolute right-0 top-full z-20 mt-1 w-40 overflow-hidden rounded-lg border border-white/10 bg-[#0a0a0a] py-1">
                <button
                  type="button"
                  className="flex w-full items-center px-3 py-2 text-left text-sm text-white/80 hover:bg-white/5"
                  onClick={onOpenSearch}
                >
                  统一搜索
                </button>
                <button
                  type="button"
                  className="flex w-full items-center px-3 py-2 text-left text-sm text-white/80 hover:bg-white/5"
                  onClick={onOpenRules}
                >
                  规则
                </button>
                <button
                  type="button"
                  className="flex w-full items-center px-3 py-2 text-left text-sm text-white/80 hover:bg-white/5"
                  onClick={onOpenOps}
                >
                  同步中心
                </button>
              </div>
            ) : null}
          </div>
        </div>

        {selectedMessageIds.length > 0 ? (
          <div className="flex flex-wrap items-center gap-2 rounded-lg border border-white/10 bg-[#0a0a0a] px-3 py-2 text-sm">
            <span className="mr-2 text-white/70">{selectedMessageIds.length} 已选</span>
            <Button variant="secondary" size="sm" onClick={onClearSelection}>
              <X className="mr-2 h-4 w-4" />
              {labels.clearSelection}
            </Button>
            <Button variant="secondary" size="sm" onClick={onMarkRead}>
              {labels.markRead}
            </Button>
            <Button variant="secondary" size="sm" onClick={onMarkUnread}>
              {labels.markUnread}
            </Button>
            <Button variant="secondary" size="sm" onClick={onStar}>
              {labels.star}
            </Button>
            <Button variant="secondary" size="sm" onClick={onArchive}>
              {labels.archive}
            </Button>
            <Button variant="secondary" size="sm" onClick={onMove}>
              {labels.move}
            </Button>
            <Button variant="destructive" size="sm" onClick={onDelete}>
              {labels.delete}
            </Button>
          </div>
        ) : null}

        {latestJob ? (
          <div className="flex flex-wrap items-center gap-2 text-[11px] text-white/45">
            <span>同步：{latestJob.status}</span>
            {latestJob.processed_messages != null ? <span>处理 {latestJob.processed_messages} 封</span> : null}
            {latestJob.finished_at ? <span>完成于 {formatMailDate(latestJob.finished_at, 'MM-dd HH:mm')}</span> : null}
          </div>
        ) : null}
      </div>

      <div className="relative min-h-0 flex-1 overflow-y-auto">
        {loadingMessages && messages.length === 0 ? (
          <div className="flex h-full items-center justify-center gap-2 text-sm text-white/45">
            <LoaderCircle className="h-4 w-4 animate-spin" />
            {labels.loading}
          </div>
        ) : messages.length === 0 ? (
          <div className="flex h-full flex-col items-center justify-center gap-2 px-6 text-center text-sm text-white/45">
            <Mail className="h-8 w-8 text-white/35" />
            <div>{labels.noEmails}</div>
            {deferredSearchQuery || filterUnread || filterStarred || filterAttachment ? (
              <button
                type="button"
                className="text-xs text-white underline underline-offset-2"
                onClick={onClearFilters}
              >
                清除筛选条件
              </button>
            ) : null}
          </div>
        ) : (
          <div className={cn('divide-y divide-white/10', loadingMessages && 'opacity-60')}>
            {messages.map((message) => (
              <div
                key={message.id}
                onClick={() => onOpenMessage(message)}
                className={cn(
                  'cursor-pointer px-3 py-3 transition-colors hover:bg-white/5',
                  activeMessageId === message.id && 'bg-white/10',
                  !message.is_read && 'font-medium',
                )}
              >
                <div className="flex items-start gap-3">
                  <input
                    type="checkbox"
                    className="mt-1 rounded border-white/20"
                    checked={selectedMessageIds.includes(message.id)}
                    onChange={(event) => {
                      event.stopPropagation();
                      onToggleMessageSelected(message.id);
                    }}
                  />
                  <div className="min-w-0 flex-1">
                    <div className="mb-1 flex items-center justify-between gap-3">
                      <span className="truncate text-sm text-white">{message.sender}</span>
                      <span className="shrink-0 text-[11px] text-white/45">
                        {formatMailDate(message.date, 'MM-dd HH:mm')}
                      </span>
                    </div>
                    <div className="truncate text-sm text-white/85">{message.subject}</div>
                    <div className="mt-1 line-clamp-2 text-xs text-white/45">{message.preview}</div>
                    <div className="mt-2 flex items-center gap-2">
                      {message.is_flagged ? <Star className="h-3.5 w-3.5 fill-amber-400 text-amber-400" /> : null}
                      {message.has_attachments ? <Paperclip className="h-3.5 w-3.5 text-white/40" /> : null}
                      {message.importance === 'high' ? <AlertCircle className="h-3.5 w-3.5 text-red-500" /> : null}
                      {(message.meta?.tags?.length ?? 0) > 0 ? (
                        <Badge variant="outline" className="max-w-[9rem] truncate">
                          {message.meta?.tags?.join(', ')}
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

      <div className="flex items-center justify-between border-t border-white/10 px-3 py-2 text-xs text-white/45">
        <label className="flex items-center gap-2">
          <input
            type="checkbox"
            className="rounded border-white/20"
            checked={allVisibleSelected}
            onChange={(event) => onToggleSelectAll(event.target.checked)}
          />
          {labels.selectAll}
        </label>
        <div className="flex items-center gap-2">
          <span className="tabular-nums">
            {listMeta.total > 0
              ? `${labels.page} ${listMeta.page || 1}/${Math.max(listMeta.total_pages || 1, 1)} · ${listMeta.total} 封`
              : `${labels.page} ${listMeta.page || 1}`}
          </span>
          <Button
            variant="ghost"
            size="icon"
            className="h-7 w-7"
            disabled={!listMeta.has_prev || loadingMessages}
            title="上一页"
            onClick={onPrevPage}
          >
            <ChevronLeft className="h-4 w-4" />
          </Button>
          <Button
            variant="ghost"
            size="icon"
            className="h-7 w-7"
            disabled={!listMeta.has_next || loadingMessages}
            title="下一页"
            onClick={onNextPage}
          >
            <ChevronRight className="h-4 w-4" />
          </Button>
        </div>
      </div>
    </div>
  );
}
