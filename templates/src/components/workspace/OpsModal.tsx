import { LoaderCircle, RefreshCw, X } from 'lucide-react';
import type { AuditLogRecord, SyncStatusRecord } from '../../lib/api';
import { methodLabel } from '../../lib/mail';
import { Badge } from '../ui/Badge';
import { Button } from '../ui/Button';

interface OpsModalProps {
  loading: boolean;
  loadingLabel: string;
  busyAction: string;
  syncStatus: SyncStatusRecord | null;
  auditLogs: AuditLogRecord[];
  onSync: () => void;
  onClose: () => void;
}

export function OpsModal({ loading, loadingLabel, busyAction, syncStatus, auditLogs, onSync, onClose }: OpsModalProps) {
  return (
    <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/70 px-4 py-6">
      <div className="flex max-h-full w-full max-w-5xl flex-col rounded-2xl border border-white/10 bg-[#0a0a0a] shadow-xl">
        <div className="flex items-center justify-between border-b border-white/10 px-5 py-4">
          <div>
            <h2 className="text-lg font-semibold">审计日志和同步中心</h2>
            <div className="mt-1 text-xs text-white/45">查看同步状态、任务历史和高频后台操作审计。</div>
          </div>
          <div className="flex items-center gap-2">
            <Button onClick={onSync} disabled={busyAction === 'sync'}>
              {busyAction === 'sync' ? <LoaderCircle className="mr-2 h-4 w-4 animate-spin" /> : <RefreshCw className="mr-2 h-4 w-4" />}
              立即同步
            </Button>
            <Button variant="ghost" size="icon" onClick={onClose}>
              <X className="h-4 w-4" />
            </Button>
          </div>
        </div>
        <div className="grid min-h-0 flex-1 gap-4 overflow-hidden px-5 py-4 lg:grid-cols-2">
          <div className="min-h-0 overflow-y-auto rounded-xl border border-white/10 p-4">
            <div className="mb-3 flex items-center justify-between">
              <h3 className="text-sm font-semibold">同步中心</h3>
              <Badge variant="outline">{syncStatus?.states.length || 0}</Badge>
            </div>
            {loading ? (
              <div className="flex items-center gap-2 text-sm text-white/45">
                <LoaderCircle className="h-4 w-4 animate-spin" />
                {loadingLabel}
              </div>
            ) : !syncStatus ? (
              <div className="text-sm text-white/45">暂无同步数据</div>
            ) : (
              <div className="space-y-3">
                {syncStatus.jobs.map((job) => (
                  <div key={job.id} className="rounded-lg border border-white/10 p-3">
                    <div className="flex items-center justify-between gap-3">
                      <span className="text-sm font-medium">Job #{job.id}</span>
                      <Badge variant={job.status === 'completed' ? 'secondary' : job.status === 'failed' ? 'destructive' : 'outline'}>
                        {job.status}
                      </Badge>
                    </div>
                    <div className="mt-2 text-xs text-white/45">
                      <div>folders: {job.folders_synced}</div>
                      <div>cached: {job.cached_messages}</div>
                      <div>{job.started_at || '-'}</div>
                    </div>
                  </div>
                ))}
                {syncStatus.states.map((state) => (
                  <div key={`${state.method}-${state.folder_id}`} className="rounded-lg border border-white/10 p-3">
                    <div className="flex items-center justify-between gap-3">
                      <span className="text-sm font-medium">{state.folder_id || 'INBOX'}</span>
                      <Badge variant={state.status === 'completed' ? 'secondary' : state.status === 'failed' ? 'destructive' : 'outline'}>
                        {state.status}
                      </Badge>
                    </div>
                    <div className="mt-2 text-xs text-white/45">
                      <div>{methodLabel(state.method)}</div>
                      <div>cached: {state.cached_messages}</div>
                      <div>{state.last_synced_at || '-'}</div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div className="min-h-0 overflow-y-auto rounded-xl border border-white/10 p-4">
            <div className="mb-3 flex items-center justify-between">
              <h3 className="text-sm font-semibold">审计日志</h3>
              <Badge variant="outline">{auditLogs.length}</Badge>
            </div>
            {auditLogs.length === 0 ? (
              <div className="text-sm text-white/45">暂无审计日志</div>
            ) : (
              <div className="space-y-3">
                {auditLogs.map((item) => (
                  <div key={item.id} className="rounded-lg border border-white/10 p-3">
                    <div className="flex items-center justify-between gap-3">
                      <div className="text-sm font-medium">{item.action}</div>
                      <Badge variant={item.status === 'success' ? 'secondary' : item.status === 'failed' ? 'destructive' : 'outline'}>
                        {item.status}
                      </Badge>
                    </div>
                    <div className="mt-2 text-xs text-white/45">
                      <div>
                        {item.target_type} · {item.target_id || '-'}
                      </div>
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
  );
}
