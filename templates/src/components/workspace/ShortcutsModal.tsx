import { X } from 'lucide-react';
import { Button } from '../ui/Button';

const SHORTCUTS: Array<[string, string]> = [
  ['j / k', '下一封 / 上一封'],
  ['c', '复制验证码'],
  ['e', '归档当前邮件'],
  ['Delete', '删除当前邮件'],
  ['/', '聚焦列表搜索'],
  ['n', '写新邮件'],
  ['r', '刷新'],
  ['u', '切换未读筛选'],
  ['Esc', '关闭面板 / 返回列表'],
];

export function ShortcutsModal({ onClose }: { onClose: () => void }) {
  return (
    <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/70 px-4 py-6">
      <div className="w-full max-w-md rounded-2xl border border-white/10 bg-[#0a0a0a] p-5 shadow-xl">
        <div className="mb-4 flex items-center justify-between">
          <h2 className="text-lg font-semibold">键盘快捷键</h2>
          <Button variant="ghost" size="icon" onClick={onClose}>
            <X className="h-4 w-4" />
          </Button>
        </div>
        <div className="space-y-2 text-sm text-white/70">
          {SHORTCUTS.map(([key, desc]) => (
            <div key={key} className="flex items-center justify-between rounded-lg border border-white/10 px-3 py-2">
              <span>{desc}</span>
              <kbd className="rounded bg-white/10 px-2 py-0.5 font-mono text-xs">{key}</kbd>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
