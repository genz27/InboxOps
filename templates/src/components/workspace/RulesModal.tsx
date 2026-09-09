import type { Dispatch, SetStateAction } from 'react';
import { LoaderCircle, ShieldCheck, X } from 'lucide-react';
import type { RuleRecord } from '../../lib/api';
import { Badge } from '../ui/Badge';
import { Button } from '../ui/Button';
import { Input } from '../ui/Input';
import { EMPTY_RULE_EDITOR, type RuleEditorState } from './types';

interface RulesModalProps {
  rules: RuleRecord[];
  loading: boolean;
  loadingLabel: string;
  busyAction: string;
  editor: RuleEditorState;
  onEditorChange: Dispatch<SetStateAction<RuleEditorState>>;
  onClose: () => void;
  onCreate: () => void;
  onApplyAll: () => void;
  onEdit: (rule: RuleRecord) => void;
  onApply: (ruleId: number) => void;
  onDelete: (ruleId: number) => void;
  onSave: () => void;
}

export function RulesModal({
  rules,
  loading,
  loadingLabel,
  busyAction,
  editor,
  onEditorChange,
  onClose,
  onCreate,
  onApplyAll,
  onEdit,
  onApply,
  onDelete,
  onSave,
}: RulesModalProps) {
  return (
    <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/70 px-4 py-6">
      <div className="flex max-h-full w-full max-w-6xl flex-col rounded-2xl border border-white/10 bg-[#0a0a0a] shadow-xl">
        <div className="flex items-center justify-between border-b border-white/10 px-5 py-4">
          <div>
            <h2 className="text-lg font-semibold">规则引擎</h2>
            <div className="mt-1 text-xs text-white/45">支持条件匹配、标签追加、标记已读、移动文件夹、跟进与备注追加。</div>
          </div>
          <Button variant="ghost" size="icon" onClick={onClose}>
            <X className="h-4 w-4" />
          </Button>
        </div>
        <div className="grid min-h-0 flex-1 gap-4 overflow-hidden px-5 py-4 lg:grid-cols-[minmax(0,1fr)_minmax(360px,420px)]">
          <div className="min-h-0 overflow-y-auto rounded-xl border border-white/10">
            <div className="flex items-center justify-between border-b border-white/10 px-4 py-3">
              <div className="flex items-center gap-2">
                <ShieldCheck className="h-4 w-4" />
                <span className="text-sm font-semibold">已保存规则</span>
              </div>
              <div className="flex gap-2">
                <Button variant="outline" size="sm" onClick={onCreate}>
                  新建规则
                </Button>
                <Button variant="outline" size="sm" onClick={onApplyAll} disabled={busyAction === 'rule-apply'}>
                  一键应用
                </Button>
              </div>
            </div>
            {loading ? (
              <div className="flex items-center justify-center gap-2 px-4 py-10 text-sm text-white/45">
                <LoaderCircle className="h-4 w-4 animate-spin" />
                {loadingLabel}
              </div>
            ) : rules.length === 0 ? (
              <div className="px-4 py-10 text-center text-sm text-white/45">当前邮箱还没有规则</div>
            ) : (
              <div className="space-y-3 p-4">
                {rules.map((rule) => (
                  <div key={rule.id} className="rounded-xl border border-white/10 p-4">
                    <div className="flex items-start justify-between gap-3">
                      <div>
                        <div className="text-sm font-semibold">{rule.name}</div>
                        <div className="mt-1 text-xs text-white/45">priority: {rule.priority}</div>
                      </div>
                      <Badge variant={rule.enabled ? 'secondary' : 'outline'}>{rule.enabled ? 'enabled' : 'disabled'}</Badge>
                    </div>
                    <pre className="mt-3 whitespace-pre-wrap rounded-lg bg-white/5 p-3 text-xs text-white/70">
                      {JSON.stringify(rule.conditions, null, 2)}
                    </pre>
                    <pre className="mt-3 whitespace-pre-wrap rounded-lg bg-white/5 p-3 text-xs text-white/70">
                      {JSON.stringify(rule.actions, null, 2)}
                    </pre>
                    <div className="mt-3 flex flex-wrap gap-2">
                      <Button variant="ghost" size="sm" onClick={() => onEdit(rule)}>
                        编辑
                      </Button>
                      <Button variant="ghost" size="sm" onClick={() => onApply(rule.id)}>
                        应用当前规则
                      </Button>
                      <Button variant="destructive" size="sm" onClick={() => onDelete(rule.id)}>
                        删除
                      </Button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div className="min-h-0 overflow-y-auto rounded-xl border border-white/10 p-4">
            <h3 className="mb-4 text-sm font-semibold">{editor.id ? '编辑规则' : '新建规则'}</h3>
            <div className="grid gap-3">
              <Input
                value={editor.name}
                onChange={(event) => onEditorChange((current) => ({ ...current, name: event.target.value }))}
                placeholder="规则名称"
              />
              <div className="grid grid-cols-[1fr_120px] gap-3">
                <label className="flex items-center gap-2 rounded-md border border-white/10 px-3 py-2 text-sm">
                  <input
                    type="checkbox"
                    checked={editor.enabled}
                    onChange={(event) => onEditorChange((current) => ({ ...current, enabled: event.target.checked }))}
                  />
                  启用规则
                </label>
                <Input
                  type="number"
                  value={editor.priority}
                  onChange={(event) => onEditorChange((current) => ({ ...current, priority: Number(event.target.value) || 100 }))}
                />
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-white/45">条件 JSON</label>
                <textarea
                  className="min-h-40 w-full rounded-md border border-white/10 bg-transparent px-3 py-2 font-mono text-sm shadow-sm"
                  value={editor.conditionsText}
                  onChange={(event) => onEditorChange((current) => ({ ...current, conditionsText: event.target.value }))}
                />
              </div>
              <div>
                <label className="mb-1 block text-xs font-medium text-white/45">动作 JSON</label>
                <textarea
                  className="min-h-40 w-full rounded-md border border-white/10 bg-transparent px-3 py-2 font-mono text-sm shadow-sm"
                  value={editor.actionsText}
                  onChange={(event) => onEditorChange((current) => ({ ...current, actionsText: event.target.value }))}
                />
              </div>
              <div className="flex justify-end gap-2">
                <Button variant="outline" onClick={() => onEditorChange(EMPTY_RULE_EDITOR)}>
                  清空
                </Button>
                <Button onClick={onSave} disabled={busyAction === 'rule-save'}>
                  {busyAction === 'rule-save' ? <LoaderCircle className="mr-2 h-4 w-4 animate-spin" /> : null}
                  保存规则
                </Button>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
