import {
  Archive,
  ChevronLeft,
  ChevronRight,
  Clock,
  Download,
  FilePlus2,
  Forward,
  LoaderCircle,
  Mail,
  MailOpen,
  Paperclip,
  Reply,
  ReplyAll,
  Star,
  Trash2,
} from 'lucide-react';
import { SafeHtml } from '../SafeHtml';
import { Badge } from '../ui/Badge';
import { Button } from '../ui/Button';
import { Input } from '../ui/Input';
import { cn } from '../../lib/utils';
import { methodLabel } from '../../lib/mail';
import type { Email, EmailAccount, MethodValue } from '../../store/useAppStore';
import { formatMailDate, isDraftFolder, type MetaFormState, type ViewMode } from './types';

export interface MessageReaderCapabilities {
  canSaveDraft: boolean;
  canReply: boolean;
  canForward: boolean;
  writeUnsupportedHint: string;
}

interface MessageReaderLabels {
  loading: string;
  markUnread: string;
  markRead: string;
  unstar: string;
  star: string;
  archive: string;
  delete: string;
  to: string;
  cc: string;
  htmlView: string;
  textView: string;
  headersView: string;
  attachments: string;
  messageId: string;
  conversationId: string;
  selectEmail: string;
}

interface MessageReaderProps {
  labels: MessageReaderLabels;
  loadingDetail: boolean;
  activeMessage: Email | null;
  activeMessageId: string;
  activeFolderId: string;
  activeMailbox: EmailAccount | null;
  activeMethod: MethodValue;
  capabilities: MessageReaderCapabilities;
  verificationCodes: string[];
  copiedCode: string;
  viewMode: ViewMode;
  detailSidebarCollapsed: boolean;
  metaPanelCollapsed: boolean;
  threadPanelCollapsed: boolean;
  metaForm: MetaFormState;
  busyAction: string;
  threadItems: Email[];
  onCloseMessage: () => void;
  onEditDraft: (message: Email) => void;
  onReply: (message: Email) => void;
  onReplyAll: (message: Email) => void;
  onForward: (message: Email) => void;
  onToggleRead: (isRead: boolean) => void;
  onToggleFlag: (isFlagged: boolean) => void;
  onArchive: (messageId: string) => void;
  onDelete: () => void;
  onCopyVerificationCode: (code: string) => void;
  onViewModeChange: (mode: ViewMode) => void;
  onToggleDetailSidebar: () => void;
  onToggleMetaPanel: () => void;
  onToggleThreadPanel: () => void;
  onMetaFormChange: (patch: Partial<MetaFormState>) => void;
  onSaveMeta: () => void;
  onOpenThreadMessage: (message: Email) => void;
  onDownloadAttachment: (attachmentId: string, filename: string) => void;
}

export function MessageReader({
  labels,
  loadingDetail,
  activeMessage,
  activeMessageId,
  activeFolderId,
  activeMailbox,
  activeMethod,
  capabilities,
  verificationCodes,
  copiedCode,
  viewMode,
  detailSidebarCollapsed,
  metaPanelCollapsed,
  threadPanelCollapsed,
  metaForm,
  busyAction,
  threadItems,
  onCloseMessage,
  onEditDraft,
  onReply,
  onReplyAll,
  onForward,
  onToggleRead,
  onToggleFlag,
  onArchive,
  onDelete,
  onCopyVerificationCode,
  onViewModeChange,
  onToggleDetailSidebar,
  onToggleMetaPanel,
  onToggleThreadPanel,
  onMetaFormChange,
  onSaveMeta,
  onOpenThreadMessage,
  onDownloadAttachment,
}: MessageReaderProps) {
  return (
    <div className={cn('min-w-0 flex-1 bg-black', !activeMessageId ? 'hidden md:flex md:flex-col' : 'flex flex-col')}>
      {loadingDetail ? (
        <div className="flex flex-1 items-center justify-center gap-2 text-sm text-white/45">
          <LoaderCircle className="h-4 w-4 animate-spin" />
          {labels.loading}
        </div>
      ) : activeMessage ? (
        <>
          <div className="flex shrink-0 flex-wrap items-center justify-between gap-3 border-b border-white/10 px-4 py-3">
            <div className="flex flex-wrap items-center gap-2">
              <Button variant="ghost" size="icon" className="md:hidden" onClick={onCloseMessage}>
                <ChevronLeft className="h-4 w-4" />
              </Button>
              {isDraftFolder(activeMessage, activeFolderId) ? (
                <Button
                  variant="outline"
                  size="icon"
                  className="h-8 w-8"
                  disabled={!capabilities.canSaveDraft}
                  title={capabilities.canSaveDraft ? '编辑草稿' : capabilities.writeUnsupportedHint}
                  onClick={() => onEditDraft(activeMessage)}
                >
                  <FilePlus2 className="h-4 w-4" />
                </Button>
              ) : null}
              <Button
                variant="ghost"
                size="icon"
                className="h-8 w-8"
                disabled={!capabilities.canReply}
                title={capabilities.canReply ? '回复' : capabilities.writeUnsupportedHint}
                onClick={() => onReply(activeMessage)}
              >
                <Reply className="h-4 w-4" />
              </Button>
              <Button
                variant="ghost"
                size="icon"
                className="h-8 w-8"
                disabled={!capabilities.canReply}
                title={capabilities.canReply ? '回复全部' : capabilities.writeUnsupportedHint}
                onClick={() => onReplyAll(activeMessage)}
              >
                <ReplyAll className="h-4 w-4" />
              </Button>
              <Button
                variant="ghost"
                size="icon"
                className="h-8 w-8"
                disabled={!capabilities.canForward}
                title={capabilities.canForward ? '转发' : capabilities.writeUnsupportedHint}
                onClick={() => onForward(activeMessage)}
              >
                <Forward className="h-4 w-4" />
              </Button>
              <Button
                variant="ghost"
                size="icon"
                className="h-8 w-8"
                title={activeMessage.is_read ? labels.markUnread : labels.markRead}
                onClick={() => onToggleRead(!activeMessage.is_read)}
              >
                {activeMessage.is_read ? <Mail className="h-4 w-4" /> : <MailOpen className="h-4 w-4" />}
              </Button>
              <Button
                variant="ghost"
                size="icon"
                className="h-8 w-8"
                title={activeMessage.is_flagged ? labels.unstar : labels.star}
                onClick={() => onToggleFlag(!activeMessage.is_flagged)}
              >
                <Star className={cn('h-4 w-4', activeMessage.is_flagged && 'fill-amber-400 text-amber-400')} />
              </Button>
              <Button
                variant="ghost"
                size="icon"
                className="h-8 w-8"
                title={`${labels.archive} (e)`}
                onClick={() => onArchive(activeMessage.id)}
              >
                <Archive className="h-4 w-4" />
              </Button>
              <Button
                variant="destructive"
                size="icon"
                className="h-8 w-8"
                title={`${labels.delete} (Delete)`}
                onClick={onDelete}
              >
                <Trash2 className="h-4 w-4" />
              </Button>
            </div>
            <div className="text-xs text-white/45">{methodLabel(activeMethod)}</div>
          </div>

          <div className="min-h-0 flex-1 overflow-y-auto px-6 py-5">
            <div className="mb-6">
              <div className="mb-2 flex items-start justify-between gap-4">
                <h1 className="text-2xl font-semibold text-white">{activeMessage.subject}</h1>
                <div className="flex items-center gap-1 text-sm text-white/45">
                  <Clock className="h-3.5 w-3.5" />
                  {formatMailDate(activeMessage.date)}
                </div>
              </div>
              <div className="space-y-1 text-sm text-white/70">
                <div>
                  <span className="font-medium">{activeMessage.sender}</span>
                  <span className="ml-2 text-white/40">{activeMessage.mailboxEmail || activeMailbox?.email}</span>
                </div>
                <div>
                  {labels.to}: {activeMessage.to_recipients.join(', ') || '-'}
                </div>
                {activeMessage.cc_recipients.length > 0 ? (
                  <div>
                    {labels.cc}: {activeMessage.cc_recipients.join(', ')}
                  </div>
                ) : null}
              </div>
              {verificationCodes.length > 0 ? (
                <div className="mt-4 rounded-xl border border-amber-500/30 bg-amber-500/10 p-3">
                  <div className="mb-2 text-xs font-semibold uppercase tracking-wide text-amber-200">
                    检测到验证码
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {verificationCodes.map((code) => (
                      <Button
                        key={code}
                        size="sm"
                        variant={copiedCode === code ? 'secondary' : 'outline'}
                        className="font-mono tracking-widest"
                        onClick={() => void onCopyVerificationCode(code)}
                        title="一键复制验证码"
                      >
                        {copiedCode === code ? `已复制 ${code}` : code}
                      </Button>
                    ))}
                  </div>
                </div>
              ) : null}
              {loadingDetail ? (
                <div className="mt-3 flex items-center gap-2 text-xs text-white/45">
                  <LoaderCircle className="h-3.5 w-3.5 animate-spin" />
                  正在加载完整正文…
                </div>
              ) : null}
            </div>

            <div className={cn('mb-6 grid gap-4', !detailSidebarCollapsed && 'xl:grid-cols-[minmax(0,1.3fr)_minmax(320px,0.9fr)]')}>
              <div className="rounded-xl border border-white/10 p-4">
                <div className="mb-3 flex items-center justify-between">
                  <h3 className="text-sm font-semibold">消息视图</h3>
                  <div className="flex items-center gap-1">
                    <Button variant={viewMode === 'html' ? 'secondary' : 'ghost'} size="sm" onClick={() => onViewModeChange('html')}>
                      {labels.htmlView}
                    </Button>
                    <Button variant={viewMode === 'text' ? 'secondary' : 'ghost'} size="sm" onClick={() => onViewModeChange('text')}>
                      {labels.textView}
                    </Button>
                    <Button variant={viewMode === 'headers' ? 'secondary' : 'ghost'} size="sm" onClick={() => onViewModeChange('headers')}>
                      {labels.headersView}
                    </Button>
                    <Button
                      variant={detailSidebarCollapsed ? 'ghost' : 'secondary'}
                      size="sm"
                      onClick={onToggleDetailSidebar}
                      title={detailSidebarCollapsed ? '展开侧边信息' : '收起侧边信息'}
                    >
                      {detailSidebarCollapsed ? <ChevronLeft className="mr-2 h-4 w-4" /> : <ChevronRight className="mr-2 h-4 w-4" />}
                      侧边信息
                    </Button>
                  </div>
                </div>

                {viewMode === 'html' ? (
                  activeMessage.body_html ? (
                    <SafeHtml html={activeMessage.body_html} minHeight={320} />
                  ) : (
                    <pre className="whitespace-pre-wrap text-sm text-white/75">{activeMessage.body_text}</pre>
                  )
                ) : null}
                {viewMode === 'text' ? (
                  <pre className="whitespace-pre-wrap text-sm text-white/75">{activeMessage.body_text}</pre>
                ) : null}
                {viewMode === 'headers' ? (
                  <pre className="whitespace-pre-wrap rounded-lg bg-white/5 p-3 text-xs text-white/70">
                    {activeMessage.headers || '(无 Headers)'}
                  </pre>
                ) : null}

                {activeMessage.attachments.length > 0 ? (
                  <div className="mt-6 border-t border-white/10 pt-4">
                    <h3 className="mb-3 flex items-center gap-2 text-sm font-semibold">
                      <Paperclip className="h-4 w-4" />
                      {labels.attachments} ({activeMessage.attachments.length})
                    </h3>
                    <div className="space-y-2">
                      {activeMessage.attachments.map((attachment) => (
                        <div
                          key={attachment.id}
                          className="flex items-center justify-between gap-3 rounded-lg border border-white/10 px-3 py-2"
                        >
                          <div className="min-w-0">
                            <div className="truncate text-sm font-medium">{attachment.name}</div>
                            <div className="text-xs text-white/45">{(attachment.size / 1024).toFixed(1)} KB</div>
                          </div>
                          <Button
                            variant="ghost"
                            size="icon"
                            className="h-8 w-8"
                            disabled={busyAction === `download-${attachment.id}`}
                            onClick={() => void onDownloadAttachment(attachment.id, attachment.name)}
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
                  <div className="rounded-xl border border-white/10 p-4">
                    <div className={cn('flex items-center justify-between gap-2', metaPanelCollapsed ? 'mb-0' : 'mb-3')}>
                      <button
                        type="button"
                        className="flex min-w-0 items-center gap-2 rounded-md text-left transition-colors hover:text-white"
                        onClick={onToggleMetaPanel}
                        aria-expanded={!metaPanelCollapsed}
                        aria-controls="workspace-meta-panel"
                      >
                        <ChevronRight
                          className={cn(
                            'h-4 w-4 shrink-0 text-white/40 transition-transform duration-200',
                            !metaPanelCollapsed && 'rotate-90',
                          )}
                        />
                        <h3 className="text-sm font-semibold">标签 / 跟进 / 备注 / Snooze</h3>
                      </button>
                      <Badge variant="outline">{activeMessage.meta?.status || 'active'}</Badge>
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
                            <label className="mb-1 block text-xs font-medium text-white/45">标签</label>
                            <Input
                              value={metaForm.tags}
                              onChange={(event) => onMetaFormChange({ tags: event.target.value })}
                              placeholder="vip, follow-up"
                            />
                          </div>
                          <div>
                            <label className="mb-1 block text-xs font-medium text-white/45">跟进</label>
                            <Input
                              value={metaForm.followUp}
                              onChange={(event) => onMetaFormChange({ followUp: event.target.value })}
                              placeholder="today / tomorrow / custom"
                            />
                          </div>
                          <div>
                            <label className="mb-1 block text-xs font-medium text-white/45">稍后提醒</label>
                            <Input
                              type="datetime-local"
                              value={metaForm.snoozedUntil}
                              onChange={(event) => onMetaFormChange({ snoozedUntil: event.target.value })}
                            />
                          </div>
                          <div>
                            <label className="mb-1 block text-xs font-medium text-white/45">状态</label>
                            <select
                              className="flex h-9 w-full rounded-md border border-white/10 bg-transparent px-3 text-sm"
                              value={metaForm.status}
                              onChange={(event) => onMetaFormChange({ status: event.target.value })}
                            >
                              <option value="active">active</option>
                              <option value="snoozed">snoozed</option>
                              <option value="done">done</option>
                            </select>
                          </div>
                          <div>
                            <label className="mb-1 block text-xs font-medium text-white/45">备注</label>
                            <textarea
                              className="min-h-28 w-full rounded-md border border-white/10 bg-transparent px-3 py-2 text-sm shadow-sm"
                              value={metaForm.notes}
                              onChange={(event) => onMetaFormChange({ notes: event.target.value })}
                            />
                          </div>
                          <Button onClick={onSaveMeta} disabled={busyAction === 'meta'}>
                            {busyAction === 'meta' ? <LoaderCircle className="mr-2 h-4 w-4 animate-spin" /> : null}
                            保存 Meta
                          </Button>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="rounded-xl border border-white/10 p-4">
                    <div className={cn('flex items-center justify-between gap-2', threadPanelCollapsed ? 'mb-0' : 'mb-3')}>
                      <button
                        type="button"
                        className="flex min-w-0 items-center gap-2 rounded-md text-left transition-colors hover:text-white"
                        onClick={onToggleThreadPanel}
                        aria-expanded={!threadPanelCollapsed}
                        aria-controls="workspace-thread-panel"
                      >
                        <ChevronRight
                          className={cn(
                            'h-4 w-4 shrink-0 text-white/40 transition-transform duration-200',
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
                            <div className="text-sm text-white/45">当前会话没有已缓存的线程邮件</div>
                          ) : (
                            <div className="space-y-2">
                              {threadItems.map((item) => (
                                <button
                                  key={item.id}
                                  onClick={() => void onOpenThreadMessage(item)}
                                  className={cn(
                                    'w-full rounded-lg border px-3 py-2 text-left transition-colors',
                                    item.id === activeMessage.id
                                      ? 'border-white/20 bg-white/10'
                                      : 'border-white/10 hover:bg-white/5',
                                  )}
                                >
                                  <div className="truncate text-sm font-medium">{item.subject}</div>
                                  <div className="mt-1 text-xs text-white/45">
                                    {item.sender} · {formatMailDate(item.date, 'MM-dd HH:mm')}
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

            <div className="rounded-xl border border-white/10 p-4 text-xs text-white/45">
              <div className="flex gap-2">
                <span className="min-w-24 font-medium">{labels.messageId}:</span>
                <span className="break-all font-mono">{activeMessage.internet_message_id || activeMessage.id}</span>
              </div>
              <div className="mt-2 flex gap-2">
                <span className="min-w-24 font-medium">{labels.conversationId}:</span>
                <span className="break-all font-mono">{activeMessage.conversation_id || '-'}</span>
              </div>
            </div>
          </div>
        </>
      ) : (
        <div className="flex flex-1 flex-col items-center justify-center gap-4 text-white/45">
          <Mail className="h-12 w-12 opacity-20" />
          <p>{labels.selectEmail}</p>
        </div>
      )}
    </div>
  );
}
