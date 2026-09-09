import React, { useEffect, useMemo, useState } from 'react';
import {
  CheckCircle2,
  Plus,
  RefreshCw,
  Search,
  Settings2,
  Trash2,
  Upload,
  XCircle,
} from 'lucide-react';
import { useI18n } from '../i18n';
import { Button } from '../components/ui/Button';
import { Input } from '../components/ui/Input';
import { Badge } from '../components/ui/Badge';
import type { EmailAccount, MethodValue } from '../store/useAppStore';
import {
  batchDeleteMailboxes,
  batchTestConnections,
  batchUpdatePreferredMethod,
  createMailbox,
  getMailbox,
  importMailboxes,
  listMailboxes,
  methodLabel,
  updateMailbox,
} from '../lib/api';

type AccountFormMode = 'create' | 'edit';

interface AccountFormState {
  mode: AccountFormMode;
  id: string;
  label: string;
  email: string;
  clientId: string;
  refreshToken: string;
  preferredMethod: MethodValue;
  proxy: string;
  notes: string;
}

interface ImportFormState {
  rawText: string;
  preferredMethod: MethodValue;
  autoDetectMethod: boolean;
}

const DEFAULT_FORM: AccountFormState = {
  mode: 'create',
  id: '',
  label: '',
  email: '',
  clientId: '',
  refreshToken: '',
  preferredMethod: 'graph_api',
  proxy: '',
  notes: '',
};

const DEFAULT_IMPORT_FORM: ImportFormState = {
  rawText: '',
  preferredMethod: 'graph_api',
  autoDetectMethod: true,
};

const METHOD_OPTIONS: Array<{ value: MethodValue; label: string }> = [
  { value: 'graph_api', label: 'Graph API' },
  { value: 'imap_new', label: '新版 IMAP' },
  { value: 'imap_old', label: '旧版 IMAP' },
];

function mergeAccountStatuses(nextAccounts: EmailAccount[], currentAccounts: EmailAccount[]): EmailAccount[] {
  const statusMap = new Map(currentAccounts.map((account) => [account.id, account.status]));
  return nextAccounts.map((account) => ({
    ...account,
    status: statusMap.get(account.id) ?? account.status,
  }));
}

export default function Accounts() {
  const { t } = useI18n();
  const [accounts, setAccounts] = useState<EmailAccount[]>([]);
  const [selectedAccountIds, setSelectedAccountIds] = useState<string[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [editingAccount, setEditingAccount] = useState<AccountFormState | null>(null);
  const [importDialogOpen, setImportDialogOpen] = useState(false);
  const [importForm, setImportForm] = useState<ImportFormState>(DEFAULT_IMPORT_FORM);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [importing, setImporting] = useState(false);
  const [batchBusy, setBatchBusy] = useState('');
  const [error, setError] = useState('');
  const [importError, setImportError] = useState('');
  const [notice, setNotice] = useState('');
  const [importProbeResults, setImportProbeResults] = useState<
    Array<{
      mailbox_id?: number;
      email?: string;
      label?: string;
      success?: boolean;
      adapted?: boolean;
      preferred_method?: string;
      original_method?: string;
      message?: string;
    }>
  >([]);
  const [importResultOpen, setImportResultOpen] = useState(false);
  const [methodFilter, setMethodFilter] = useState<'all' | MethodValue>('all');
  const [batchMethod, setBatchMethod] = useState<MethodValue>('graph_api');
  const [confirmDeleteOpen, setConfirmDeleteOpen] = useState(false);

  const filteredAccounts = useMemo(
    () =>
      accounts.filter((account) => {
        if (methodFilter !== 'all') {
          const currentMethod = account.preferredMethod || account.method;
          if (currentMethod !== methodFilter) {
            return false;
          }
        }
        const needle = searchQuery.trim().toLowerCase();
        if (!needle) {
          return true;
        }
        return (
          account.email.toLowerCase().includes(needle) ||
          account.label.toLowerCase().includes(needle) ||
          account.notes.toLowerCase().includes(needle)
        );
      }),
    [accounts, searchQuery, methodFilter],
  );

  const loadAccounts = async () => {
    setLoading(true);
    setError('');
    try {
      const nextAccounts = await listMailboxes();
      setAccounts((current) => mergeAccountStatuses(nextAccounts, current));
      setSelectedAccountIds((current) => current.filter((item) => nextAccounts.some((account) => account.id === item)));
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : '加载邮箱档案失败');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void loadAccounts();
  }, []);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') {
        return;
      }
      setEditingAccount(null);
      setImportDialogOpen(false);
      setImportResultOpen(false);
      setConfirmDeleteOpen(false);
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, []);

  const openCreateModal = () => {
    setEditingAccount({ ...DEFAULT_FORM });
  };

  const openImportModal = () => {
    setImportError('');
    setImportDialogOpen(true);
  };

  const closeImportModal = () => {
    setImportDialogOpen(false);
    setImportForm(DEFAULT_IMPORT_FORM);
    setImportError('');
  };

  const openEditModal = async (account: EmailAccount) => {
    setSaving(true);
    setError('');
    try {
      const detail = await getMailbox(account.id);
      setEditingAccount({
        mode: 'edit',
        id: String(detail.id),
        label: detail.label,
        email: detail.email,
        clientId: '',
        refreshToken: '',
        preferredMethod: detail.preferred_method,
        proxy: detail.proxy ?? '',
        notes: detail.notes ?? '',
      });
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : '读取邮箱档案失败');
    } finally {
      setSaving(false);
    }
  };

  const handleToggleAccount = (accountId: string) => {
    setSelectedAccountIds((current) =>
      current.includes(accountId) ? current.filter((item) => item !== accountId) : [...current, accountId],
    );
  };

  const handleImport = async (event: React.FormEvent) => {
    event.preventDefault();
    const rawText = importForm.rawText.trim();
    if (!rawText) {
      setImportError('批量导入内容不能为空');
      return;
    }
    setImporting(true);
    setError('');
    setImportError('');
    try {
      const result = await importMailboxes(rawText, importForm.preferredMethod, {
        autoDetectMethod: importForm.autoDetectMethod,
      });
      await loadAccounts();
      closeImportModal();
      const adapted = Number(result.summary.method_adapted ?? 0);
      const failed = Number(result.summary.method_failed ?? 0);
      if (result.methodProbeResults.length > 0) {
        setImportProbeResults(
          result.methodProbeResults.map((item) => ({
            mailbox_id: typeof item.mailbox_id === 'number' ? item.mailbox_id : undefined,
            email: typeof item.email === 'string' ? item.email : '',
            label: typeof item.label === 'string' ? item.label : '',
            success: item.success === true,
            adapted: item.adapted === true,
            preferred_method: typeof item.preferred_method === 'string' ? item.preferred_method : '',
            original_method: typeof item.original_method === 'string' ? item.original_method : '',
            message: typeof item.message === 'string' ? item.message : '',
          })),
        );
        setImportResultOpen(true);
      }
      if (importForm.autoDetectMethod) {
        setNotice(
          adapted > 0 || failed > 0
            ? `导入完成。已自适应 ${adapted} 个接入方式${failed > 0 ? `，${failed} 个探测失败（保留默认方式）` : ''}。`
            : t('importSuccess'),
        );
      } else {
        setNotice(t('importSuccess'));
      }
    } catch (requestError) {
      setImportError(requestError instanceof Error ? requestError.message : t('importFailed'));
    } finally {
      setImporting(false);
    }
  };

  const handleSaveAccount = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!editingAccount) {
      return;
    }
    setSaving(true);
    setError('');
    try {
      if (editingAccount.mode === 'create') {
        await createMailbox({
          label: editingAccount.label.trim(),
          email: editingAccount.email.trim(),
          client_id: editingAccount.clientId.trim(),
          refresh_token: editingAccount.refreshToken.trim(),
          preferred_method: editingAccount.preferredMethod,
          proxy: editingAccount.proxy.trim(),
          notes: editingAccount.notes.trim(),
        });
      } else {
        const payload: Record<string, string> = {
          label: editingAccount.label.trim(),
          preferred_method: editingAccount.preferredMethod,
          proxy: editingAccount.proxy.trim(),
          notes: editingAccount.notes.trim(),
        };
        if (editingAccount.clientId.trim()) {
          payload.client_id = editingAccount.clientId.trim();
        }
        if (editingAccount.refreshToken.trim()) {
          payload.refresh_token = editingAccount.refreshToken.trim();
        }
        await updateMailbox(editingAccount.id, payload);
      }
      setEditingAccount(null);
      setNotice(editingAccount.mode === 'create' ? t('accountCreated') : t('accountUpdated'));
      await loadAccounts();
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : '保存邮箱档案失败');
    } finally {
      setSaving(false);
    }
  };

  const handleBatchDelete = async () => {
    if (selectedAccountIds.length === 0 || batchBusy) {
      return;
    }
    setConfirmDeleteOpen(false);
    setBatchBusy('delete');
    setError('');
    try {
      const payload = await batchDeleteMailboxes(selectedAccountIds);
      setNotice(`已处理 ${payload.summary.processed} 个档案，成功 ${payload.summary.succeeded} 个，失败 ${payload.summary.failed} 个。`);
      await loadAccounts();
      setSelectedAccountIds([]);
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : '批量删除失败');
    } finally {
      setBatchBusy('');
    }
  };

  const handleBatchTest = async (mailboxIds: string[]) => {
    if (mailboxIds.length === 0 || batchBusy) {
      return;
    }
    setBatchBusy('test');
    setError('');
    try {
      const payload = await batchTestConnections(mailboxIds);
      const statusMap = new Map<string, EmailAccount['status']>();
      payload.results.forEach((item) => {
        if (typeof item.mailbox_id === 'number') {
          statusMap.set(String(item.mailbox_id), item.success === true ? 'connected' : 'disconnected');
        }
      });
      setAccounts((current) =>
        current.map((account) => {
          const nextStatus = statusMap.get(account.id);
          return nextStatus ? { ...account, status: nextStatus } : account;
        }),
      );
      const firstFailure = payload.results.find((item) => item.success === false) as { message?: string } | undefined;
      setNotice(
        firstFailure?.message
          ? `连接测试完成：成功 ${payload.summary.succeeded}，失败 ${payload.summary.failed}。首个失败：${firstFailure.message}`
          : `连接测试完成：成功 ${payload.summary.succeeded} 个。`,
      );
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : '批量测试失败');
    } finally {
      setBatchBusy('');
    }
  };

  const handleBatchChangeMethod = async () => {
    if (selectedAccountIds.length === 0 || batchBusy) {
      return;
    }
    setBatchBusy('method');
    setError('');
    try {
      const payload = await batchUpdatePreferredMethod(selectedAccountIds, batchMethod);
      setNotice(`已切换 ${payload.summary.succeeded} 个档案的默认接入方式。`);
      await loadAccounts();
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : '批量切换失败');
    } finally {
      setBatchBusy('');
    }
  };

  useEffect(() => {
    if (!notice) {
      return;
    }
    const timerId = window.setTimeout(() => setNotice(''), 3500);
    return () => window.clearTimeout(timerId);
  }, [notice]);

  return (
    <div className="relative flex h-full flex-1 flex-col overflow-hidden bg-black p-6">
      <div className="mb-6 flex shrink-0 items-center justify-between">
        <div>
          <h1 className="text-xl font-medium tracking-tight text-white">{t('accounts')}</h1>
          <p className="mt-1 text-sm text-white/45">
            {t('total')}: {filteredAccounts.length}
            {filteredAccounts.length !== accounts.length ? ` / ${accounts.length}` : ''}
          </p>
        </div>
        <div className="flex items-center gap-3">
          <Button variant="outline" onClick={openImportModal}>
            <Upload className="mr-2 h-4 w-4" />
            {t('import')}
          </Button>
          <Button onClick={openCreateModal}>
            <Plus className="mr-2 h-4 w-4" />
            {t('addAccount')}
          </Button>
        </div>
      </div>

      <div className="mb-4 flex shrink-0 flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-2">
          <div className="relative w-72">
            <Search className="absolute left-3 top-2.5 h-4 w-4 text-white/35" />
            <Input
              placeholder={t('search')}
              className="pl-9"
              value={searchQuery}
              onChange={(event) => setSearchQuery(event.target.value)}
            />
          </div>
          <select
            className="flex h-9 rounded-md border border-white/15 bg-black px-3 text-sm text-white"
            value={methodFilter}
            onChange={(event) => setMethodFilter(event.target.value as 'all' | MethodValue)}
          >
            <option value="all">{t('allMethods')}</option>
            {METHOD_OPTIONS.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>

        {selectedAccountIds.length > 0 ? (
          <div className="flex flex-wrap items-center gap-2 rounded-md border border-white/10 bg-[#0a0a0a] px-3 py-1.5">
            <span className="mr-1 text-sm text-white/70">
              {selectedAccountIds.length} {t('selected')}
            </span>
            <Button variant="secondary" size="sm" onClick={() => setSelectedAccountIds([])}>
              <XCircle className="mr-2 h-4 w-4" />
              {t('clearSelection')}
            </Button>
            <Button variant="secondary" size="sm" onClick={() => void handleBatchTest(selectedAccountIds)}>
              <RefreshCw className="mr-2 h-4 w-4" />
              {t('batchTest')}
            </Button>
            <select
              className="flex h-8 rounded-md border border-white/15 bg-black px-2 text-xs text-white"
              value={batchMethod}
              onChange={(event) => setBatchMethod(event.target.value as MethodValue)}
            >
              {METHOD_OPTIONS.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
            <Button variant="secondary" size="sm" onClick={() => void handleBatchChangeMethod()}>
              <Settings2 className="mr-2 h-4 w-4" />
              {t('batchChangeMethod')}
            </Button>
            <Button variant="destructive" size="sm" onClick={() => setConfirmDeleteOpen(true)}>
              <Trash2 className="mr-2 h-4 w-4" />
              {t('batchDelete')}
            </Button>
          </div>
        ) : null}
      </div>

      {error ? (
        <div className="mb-4 flex items-start justify-between gap-2 rounded-md border border-[#ff4d4d]/30 bg-[#ff4d4d]/10 px-3 py-2 text-sm text-[#ff8080]">
          <span className="min-w-0 flex-1 break-words">{error}</span>
          <button type="button" className="shrink-0 opacity-70 hover:opacity-100" onClick={() => setError('')}>
            <XCircle className="h-4 w-4" />
          </button>
        </div>
      ) : null}
      {notice ? (
        <div className="mb-4 rounded-md border border-white/10 bg-white/5 px-3 py-2 text-sm text-white/80">
          {notice}
        </div>
      ) : null}
      {batchBusy ? (
        <div className="mb-4 text-xs text-white/40">
          {batchBusy === 'test' ? '正在测试连接…' : batchBusy === 'delete' ? '正在删除…' : '正在更新…'}
        </div>
      ) : null}

      <div className="flex flex-1 flex-col overflow-hidden rounded-lg border border-white/10 bg-[#0a0a0a]">
        <div className="flex-1 overflow-auto">
          <table className="w-full text-left text-sm">
            <thead className="sticky top-0 border-b border-white/10 bg-[#0a0a0a] text-xs uppercase tracking-wide text-white/40">
              <tr>
                <th className="w-12 p-4">
                  <input
                    type="checkbox"
                    className="rounded border-slate-300"
                    checked={
                      filteredAccounts.length > 0 &&
                      filteredAccounts.every((account) => selectedAccountIds.includes(account.id))
                    }
                    onChange={(event) => {
                      if (event.target.checked) {
                        setSelectedAccountIds((current) =>
                          Array.from(new Set([...current, ...filteredAccounts.map((account) => account.id)])),
                        );
                      } else {
                        const filteredIds = new Set(filteredAccounts.map((account) => account.id));
                        setSelectedAccountIds((current) => current.filter((id) => !filteredIds.has(id)));
                      }
                    }}
                  />
                </th>
                <th className="p-4 font-medium">{t('email')}</th>
                <th className="p-4 font-medium">{t('method')}</th>
                <th className="p-4 font-medium">{t('status')}</th>
                <th className="p-4 font-medium">{t('remark')}</th>
                <th className="p-4 text-right font-medium">{t('actions')}</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-white/10">
              {loading ? (
                <tr>
                  <td colSpan={6} className="p-8 text-center text-white/40">
                    {t('loadingAccounts')}
                  </td>
                </tr>
              ) : accounts.length === 0 ? (
                <tr>
                  <td colSpan={6} className="p-12 text-center">
                    <div className="text-sm text-white/70">{t('noAccounts')}</div>
                    <div className="mt-2 text-xs text-white/40">{t('noAccountsHint')}</div>
                    <Button className="mt-4" size="sm" onClick={openCreateModal}>
                      {t('addAccount')}
                    </Button>
                  </td>
                </tr>
              ) : filteredAccounts.length === 0 ? (
                <tr>
                  <td colSpan={6} className="p-8 text-center text-white/40">
                    {t('noMatchingAccounts')}
                  </td>
                </tr>
              ) : (
                filteredAccounts.map((account) => (
                  <tr key={account.id} className="transition duration-200 ease-[cubic-bezier(0.32,0.72,0,1)] hover:bg-white/5">
                    <td className="p-4">
                      <input
                        type="checkbox"
                        className="rounded border-slate-300"
                        checked={selectedAccountIds.includes(account.id)}
                        onChange={() => handleToggleAccount(account.id)}
                      />
                    </td>
                    <td className="p-4">
                      <div className="font-medium text-white">{account.email}</div>
                      <div className="text-xs text-white/40">{account.label}</div>
                    </td>
                    <td className="p-4">
                      <Badge variant="outline">{methodLabel(account.method)}</Badge>
                    </td>
                    <td className="p-4">
                      <div className="flex items-center gap-2">
                        {account.status === 'disconnected' ? (
                          <XCircle className="h-4 w-4 text-red-500" />
                        ) : (
                          <CheckCircle2
                            className={account.status === 'connected' ? 'h-4 w-4 text-white' : 'h-4 w-4 text-white/30'}
                          />
                        )}
                        <span
                          className={
                            account.status === 'connected'
                              ? 'text-white'
                              : account.status === 'disconnected'
                                ? 'text-[#ff8080]'
                                : 'text-white/40'
                          }
                        >
                          {t(account.status)}
                        </span>
                      </div>
                    </td>
                    <td className="p-4 text-white/45">{account.notes || '-'}</td>
                    <td className="p-4 text-right">
                      <div className="flex items-center justify-end gap-2">
                        <Button variant="ghost" size="sm" onClick={() => void handleBatchTest([account.id])}>
                          {t('testConnection')}
                        </Button>
                        <Button variant="ghost" size="sm" onClick={() => void openEditModal(account)}>
                          {t('editAccount')}
                        </Button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {editingAccount ? (
        <div
          className="absolute inset-0 z-50 flex items-center justify-center bg-black/70"
          onClick={() => setEditingAccount(null)}
        >
          <div
            className="w-full max-w-xl rounded-xl border border-white/10 bg-[#0a0a0a] p-6"
            onClick={(event) => event.stopPropagation()}
          >
            <h2 className="mb-4 text-lg font-medium tracking-tight">
              {editingAccount.mode === 'create' ? t('addAccount') : t('editAccount')}
            </h2>
            <form onSubmit={handleSaveAccount} className="space-y-4">
              <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                <div>
                  <label className="mb-1 block text-sm font-medium">{t('remark')}</label>
                  <Input
                    value={editingAccount.label}
                    onChange={(event) =>
                      setEditingAccount((current) => (current ? { ...current, label: event.target.value } : current))
                    }
                    required
                  />
                </div>
                <div>
                  <label className="mb-1 block text-sm font-medium">{t('email')}</label>
                  <Input
                    value={editingAccount.email}
                    onChange={(event) =>
                      setEditingAccount((current) => (current ? { ...current, email: event.target.value } : current))
                    }
                    disabled={editingAccount.mode === 'edit'}
                    required
                  />
                </div>
              </div>
              <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                <div>
                  <label className="mb-1 block text-sm font-medium">Client ID</label>
                  <Input
                    value={editingAccount.clientId}
                    onChange={(event) =>
                      setEditingAccount((current) => (current ? { ...current, clientId: event.target.value } : current))
                    }
                    placeholder={editingAccount.mode === 'edit' ? '留空则保持不变' : ''}
                    required={editingAccount.mode === 'create'}
                  />
                </div>
                <div>
                  <label className="mb-1 block text-sm font-medium">Refresh Token</label>
                  <Input
                    value={editingAccount.refreshToken}
                    onChange={(event) =>
                      setEditingAccount((current) =>
                        current ? { ...current, refreshToken: event.target.value } : current,
                      )
                    }
                    placeholder={editingAccount.mode === 'edit' ? '留空则保持不变' : ''}
                    required={editingAccount.mode === 'create'}
                  />
                </div>
              </div>
              <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                <div>
                  <label className="mb-1 block text-sm font-medium">{t('defaultMethod')}</label>
                  <select
                    className="flex h-9 w-full rounded-md border border-white/15 bg-black px-3 py-1 text-sm text-white"
                    value={editingAccount.preferredMethod}
                    onChange={(event) =>
                      setEditingAccount((current) =>
                        current
                          ? {
                              ...current,
                              preferredMethod: event.target.value as MethodValue,
                            }
                          : current,
                      )
                    }
                  >
                    {METHOD_OPTIONS.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="mb-1 block text-sm font-medium">{t('proxyConfig')}</label>
                  <Input
                    value={editingAccount.proxy}
                    onChange={(event) =>
                      setEditingAccount((current) => (current ? { ...current, proxy: event.target.value } : current))
                    }
                    placeholder="http://proxy:port"
                  />
                </div>
              </div>
              <div>
                <label className="mb-1 block text-sm font-medium">{t('remark')}</label>
                <textarea
                  className="min-h-24 w-full rounded-md border border-white/15 bg-black px-3 py-2 text-sm text-white"
                  value={editingAccount.notes}
                  onChange={(event) =>
                    setEditingAccount((current) => (current ? { ...current, notes: event.target.value } : current))
                  }
                />
              </div>
              <div className="mt-6 flex justify-end gap-2">
                <Button type="button" variant="outline" onClick={() => setEditingAccount(null)}>
                  {t('cancel')}
                </Button>
                <Button type="submit" disabled={saving}>
                  {saving ? '保存中...' : t('save')}
                </Button>
              </div>
            </form>
          </div>
        </div>
      ) : null}

      {importDialogOpen ? (
        <div
          className="absolute inset-0 z-50 flex items-center justify-center bg-black/70 px-4 py-6"
          onClick={closeImportModal}
        >
          <div
            className="flex max-h-full w-full max-w-2xl flex-col overflow-hidden rounded-xl border border-white/10 bg-[#0a0a0a]"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="border-b border-white/10 px-6 py-4">
              <h2 className="text-lg font-medium tracking-tight">{t('importAccounts')}</h2>
              <p className="mt-1 text-sm text-white/45">{t('importLineHint')}</p>
            </div>
            <form onSubmit={handleImport} className="flex min-h-0 flex-1 flex-col">
              <div className="min-h-0 space-y-4 overflow-y-auto px-6 py-5">
                <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                  <div>
                    <label className="mb-1 block text-sm font-medium">{t('defaultMethod')}</label>
                    <select
                      className="flex h-9 w-full rounded-md border border-white/15 bg-black px-3 py-1 text-sm text-white"
                      value={importForm.preferredMethod}
                      onChange={(event) =>
                        setImportForm((current) => ({
                          ...current,
                          preferredMethod: event.target.value as MethodValue,
                        }))
                      }
                      disabled={importForm.autoDetectMethod}
                    >
                      {METHOD_OPTIONS.map((option) => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>
                  </div>
                  <label className="flex items-end gap-2 pb-1 text-sm text-white/80">
                    <input
                      type="checkbox"
                      className="rounded border-slate-300"
                      checked={importForm.autoDetectMethod}
                      onChange={(event) =>
                        setImportForm((current) => ({
                          ...current,
                          autoDetectMethod: event.target.checked,
                        }))
                      }
                    />
                    <span>
                      导入后自动探测可用接入方式
                      <span className="mt-0.5 block text-xs text-white/40">
                        顺序：Graph → 新版 IMAP → 旧版 IMAP
                      </span>
                    </span>
                  </label>
                </div>

                <div>
                  <label className="mb-1 block text-sm font-medium">{t('importPayload')}</label>
                  <textarea
                    className="min-h-64 w-full rounded-md border border-white/15 bg-black px-3 py-2 font-mono text-sm text-white"
                    value={importForm.rawText}
                    onChange={(event) =>
                      setImportForm((current) => ({
                        ...current,
                        rawText: event.target.value,
                      }))
                    }
                    placeholder={t('importPlaceholder')}
                    spellCheck={false}
                  />
                </div>

                <div className="rounded-lg border border-dashed border-white/15 bg-black px-4 py-3 text-xs leading-6 text-white/45">
                  <div>{t('importSupportedFormats')}</div>
                  <div>{t('importExampleSimple')}</div>
                  <div>{t('importExampleKeyed')}</div>
                </div>

                {importError ? <div className="text-sm text-[#ff8080]">{importError}</div> : null}
              </div>

              <div className="flex justify-end gap-2 border-t border-white/10 px-6 py-4">
                <Button type="button" variant="outline" onClick={closeImportModal} disabled={importing}>
                  {t('cancel')}
                </Button>
                <Button type="submit" disabled={importing || !importForm.rawText.trim()}>
                  {importing ? t('importing') : t('confirmImport')}
                </Button>
              </div>
            </form>
          </div>
        </div>
      ) : null}

      {importResultOpen ? (
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/70 px-4 py-6">
          <div className="flex max-h-full w-full max-w-4xl flex-col overflow-hidden rounded-xl border border-white/10 bg-[#0a0a0a]">
            <div className="flex items-center justify-between border-b border-white/10 px-5 py-4">
              <div>
                <h2 className="text-lg font-medium tracking-tight">导入探测明细</h2>
                <p className="mt-1 text-sm text-white/45">每个账号的接入方式探测结果与最终采用方式</p>
              </div>
              <Button variant="outline" size="sm" onClick={() => setImportResultOpen(false)}>
                关闭
              </Button>
            </div>
            <div className="min-h-0 flex-1 overflow-auto">
              <table className="w-full text-left text-sm">
                <thead className="sticky top-0 border-b border-white/10 bg-[#0a0a0a] text-xs uppercase text-white/40">
                  <tr>
                    <th className="p-3">邮箱</th>
                    <th className="p-3">结果</th>
                    <th className="p-3">原方式</th>
                    <th className="p-3">最终方式</th>
                    <th className="p-3">说明</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200 dark:divide-slate-800">
                  {importProbeResults.map((item, index) => (
                    <tr key={`${item.mailbox_id ?? item.email ?? index}`}>
                      <td className="p-3">
                        <div className="font-medium">{item.email || '-'}</div>
                        <div className="text-xs text-slate-500">{item.label || ''}</div>
                      </td>
                      <td className="p-3">
                        {item.success ? (
                          <span className="text-emerald-600">{item.adapted ? '已自适应' : '可用'}</span>
                        ) : (
                          <span className="text-red-600">失败</span>
                        )}
                      </td>
                      <td className="p-3">{item.original_method || '-'}</td>
                      <td className="p-3">{item.preferred_method || '-'}</td>
                      <td className="max-w-xs break-words p-3 text-xs text-slate-500">{item.message || '-'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      ) : null}

      {confirmDeleteOpen ? (
        <div
          className="absolute inset-0 z-50 flex items-center justify-center bg-black/70 px-4"
          onClick={() => setConfirmDeleteOpen(false)}
        >
          <div
            className="w-full max-w-sm rounded-xl border border-white/10 bg-[#0a0a0a] p-5"
            onClick={(event) => event.stopPropagation()}
          >
            <h2 className="text-base font-medium tracking-tight">{t('batchDelete')}</h2>
            <p className="mt-2 text-sm text-white/55">
              {t('confirmDeleteAccounts').replace('{count}', String(selectedAccountIds.length))}
            </p>
            <div className="mt-5 flex justify-end gap-2">
              <Button type="button" variant="outline" onClick={() => setConfirmDeleteOpen(false)}>
                {t('cancel')}
              </Button>
              <Button type="button" variant="destructive" onClick={() => void handleBatchDelete()}>
                {t('batchDelete')}
              </Button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}
