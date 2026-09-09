import { LoaderCircle, X } from 'lucide-react';
import { Badge } from '../ui/Badge';
import { Button } from '../ui/Button';
import { Input } from '../ui/Input';
import type { ComposeFormState } from './types';

interface ComposeDrawerProps {
  compose: ComposeFormState;
  mailboxEmail: string;
  subjectPlaceholder: string;
  onClose: () => void;
  onChange: (patch: Partial<ComposeFormState>) => void;
  onFiles: (files: FileList | null) => void;
  onRemoveAttachment: (index: number) => void;
  onSaveDraft: () => void;
  onSend: () => void;
}

function composeTitle(mode: ComposeFormState['mode']) {
  if (mode === 'new') {
    return '写信 / 发信';
  }
  if (mode === 'reply') {
    return '回复';
  }
  if (mode === 'replyAll') {
    return '回复全部';
  }
  return '转发';
}

export function ComposeDrawer({
  compose,
  mailboxEmail,
  subjectPlaceholder,
  onClose,
  onChange,
  onFiles,
  onRemoveAttachment,
  onSaveDraft,
  onSend,
}: ComposeDrawerProps) {
  return (
    <div className="absolute inset-0 z-50 flex justify-end bg-black/70">
      <button type="button" className="h-full flex-1 cursor-default" onClick={onClose} aria-label="关闭写信面板" />
      <div className="flex h-full w-full max-w-xl flex-col border-l border-white/10 bg-[#0a0a0a] shadow-2xl">
        <div className="flex items-center justify-between border-b border-white/10 px-5 py-4">
          <div>
            <h2 className="text-lg font-semibold">{composeTitle(compose.mode)}</h2>
            <div className="mt-1 text-xs text-white/45">{mailboxEmail}</div>
          </div>
          <Button variant="ghost" size="icon" onClick={onClose}>
            <X className="h-4 w-4" />
          </Button>
        </div>
        <div className="min-h-0 flex-1 space-y-4 overflow-y-auto px-5 py-4">
          <div className="grid gap-3">
            <Input value={compose.to} onChange={(event) => onChange({ to: event.target.value })} placeholder="To" />
            <Input value={compose.cc} onChange={(event) => onChange({ cc: event.target.value })} placeholder="Cc" />
            <Input value={compose.bcc} onChange={(event) => onChange({ bcc: event.target.value })} placeholder="Bcc" />
            <Input
              value={compose.subject}
              onChange={(event) => onChange({ subject: event.target.value })}
              placeholder={subjectPlaceholder}
            />
            <textarea
              className="min-h-[16rem] w-full rounded-md border border-white/10 bg-transparent px-3 py-3 text-sm shadow-sm"
              value={compose.bodyText}
              onChange={(event) => onChange({ bodyText: event.target.value })}
            />
            <div className="rounded-lg border border-dashed border-white/15 px-3 py-3">
              <div className="mb-2 text-sm font-medium">附件上传</div>
              <input type="file" multiple onChange={(event) => onFiles(event.target.files)} />
              {compose.attachmentNames.length > 0 ? (
                <div className="mt-3 flex flex-wrap gap-2">
                  {compose.attachmentNames.map((name, index) => (
                    <Badge key={`${name}-${index}`} variant="outline" className="gap-2">
                      {name}
                      <button type="button" onClick={() => onRemoveAttachment(index)}>
                        <X className="h-3 w-3" />
                      </button>
                    </Badge>
                  ))}
                </div>
              ) : null}
            </div>
          </div>
        </div>
        <div className="flex items-center justify-between gap-3 border-t border-white/10 px-5 py-4">
          <div className="text-xs text-white/45">侧滑写信面板 · 支持草稿与回复</div>
          <div className="flex items-center gap-2">
            <Button variant="outline" onClick={onSaveDraft} disabled={compose.submitting}>
              {compose.submitting ? <LoaderCircle className="mr-2 h-4 w-4 animate-spin" /> : null}
              保存草稿
            </Button>
            <Button onClick={onSend} disabled={compose.submitting}>
              {compose.submitting ? <LoaderCircle className="mr-2 h-4 w-4 animate-spin" /> : null}
              发送
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
}
