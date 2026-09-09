import { format } from 'date-fns';
import { LoaderCircle, Search, X } from 'lucide-react';
import type { PaginationMeta } from '../../lib/api';
import type { Email } from '../../store/useAppStore';
import { Button } from '../ui/Button';
import { Input } from '../ui/Input';

interface GlobalSearchModalProps {
  query: string;
  searchAcrossAll: boolean;
  loading: boolean;
  results: Email[];
  meta: PaginationMeta;
  onQueryChange: (value: string) => void;
  onSearchAcrossAllChange: (value: boolean) => void;
  onSearch: (page?: number) => void;
  onJump: (message: Email) => void;
  onClose: () => void;
}

export function GlobalSearchModal({
  query,
  searchAcrossAll,
  loading,
  results,
  meta,
  onQueryChange,
  onSearchAcrossAllChange,
  onSearch,
  onJump,
  onClose,
}: GlobalSearchModalProps) {
  return (
    <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/70 px-4 py-6">
      <div className="flex max-h-full w-full max-w-5xl flex-col rounded-2xl border border-white/10 bg-[#0a0a0a] shadow-xl">
        <div className="flex items-center justify-between border-b border-white/10 px-5 py-4">
          <div>
            <h2 className="text-lg font-semibold">跨邮箱统一搜索</h2>
            <div className="mt-1 text-xs text-white/45">基于本地索引与缓存搜索，适合跨邮箱、标签、备注和会话定位。</div>
          </div>
          <Button variant="ghost" size="icon" onClick={onClose}>
            <X className="h-4 w-4" />
          </Button>
        </div>
        <div className="space-y-4 px-5 py-4">
          <div className="flex flex-wrap items-center gap-3">
            <div className="min-w-[280px] flex-1">
              <Input
                value={query}
                onChange={(event) => onQueryChange(event.target.value)}
                placeholder="输入主题、发件人、备注或标签关键词"
              />
            </div>
            <label className="flex items-center gap-2 text-sm">
              <input type="checkbox" checked={!searchAcrossAll} onChange={(event) => onSearchAcrossAllChange(!event.target.checked)} />
              仅当前邮箱
            </label>
            <Button onClick={() => onSearch(1)} disabled={loading}>
              {loading ? <LoaderCircle className="mr-2 h-4 w-4 animate-spin" /> : <Search className="mr-2 h-4 w-4" />}
              搜索
            </Button>
          </div>
          <div className="min-h-[22rem] overflow-y-auto rounded-xl border border-white/10">
            {results.length === 0 ? (
              <div className="flex h-full items-center justify-center px-6 py-12 text-sm text-white/45">
                {loading ? '正在搜索...' : '暂无搜索结果'}
              </div>
            ) : (
              <div className="divide-y divide-white/10">
                {results.map((item) => (
                  <button key={`${item.mailboxId}-${item.id}`} onClick={() => onJump(item)} className="w-full px-4 py-3 text-left hover:bg-white/5">
                    <div className="flex items-center justify-between gap-3">
                      <div className="min-w-0">
                        <div className="truncate text-sm font-medium">{item.subject}</div>
                        <div className="mt-1 truncate text-xs text-white/45">
                          {item.mailboxEmail || '-'} · {item.sender} · {item.folderId}
                        </div>
                      </div>
                      <div className="text-[11px] text-white/45">{item.date ? format(new Date(item.date), 'MM-dd HH:mm') : '-'}</div>
                    </div>
                    {item.meta?.tags.length ? <div className="mt-2 text-xs text-white/45">标签: {item.meta.tags.join(', ')}</div> : null}
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="flex items-center justify-end gap-2">
            <Button variant="ghost" size="sm" disabled={!meta.has_prev} onClick={() => onSearch(Math.max(1, meta.page - 1))}>
              上一页
            </Button>
            <span className="text-xs text-white/45">
              {meta.page || 1} / {meta.total_pages || 1}
            </span>
            <Button variant="ghost" size="sm" disabled={!meta.has_next} onClick={() => onSearch(meta.page + 1)}>
              下一页
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
}
