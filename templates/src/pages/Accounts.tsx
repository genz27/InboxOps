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
  const [error, setError] = useState('');
  const [importError, setImportError] = useState('');

  const filteredAccounts = useMemo(
    () =>
      accounts.filter((account) => {
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
    [accounts, searchQuery],
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
      await importMailboxes(rawText, importForm.preferredMethod);
      await loadAccounts();
      closeImportModal();
      window.alert(t('importSuccess'));
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
      await loadAccounts();
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : '保存邮箱档案失败');
    } finally {
      setSaving(false);
    }
  };

  const handleBatchDelete = async () => {
    if (selectedAccountIds.length === 0) {
      return;
    }
    if (!window.confirm(`确认删除 ${selectedAccountIds.length} 个邮箱档案吗？`)) {
      return;
    }
    try {
      const payload = await batchDeleteMailboxes(selectedAccountIds);
      window.alert(`已处理 ${payload.summary.processed} 个档案，成功 ${payload.summary.succeeded} 个。`);
      await loadAccounts();
      setSelectedAccountIds([]);
    } catch (requestError) {
      window.alert(requestError instanceof Error ? requestError.message : '批量删除失败');
    }
  };

  const handleBatchTest = async (mailboxIds: string[]) => {
    if (mailboxIds.length === 0) {
      return;
    }
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
      window.alert(
        firstFailure?.message
          ? `连接测试完成，成功 ${payload.summary.succeeded} 个，失败 ${payload.summary.failed} 个。首个失败：${firstFailure.message}`
          : `连接测试完成，成功 ${payload.summary.succeeded} 个。`,
      );
    } catch (requestError) {
      window.alert(requestError instanceof Error ? requestError.message : '批量测试失败');
    }
  };

  const handleBatchChangeMethod = async () => {
    if (selectedAccountIds.length === 0) {
      return;
    }
    const nextMethod = window.prompt('请输入新的默认方式：graph_api / imap_new / imap_old', 'graph_api') as
      | MethodValue
      | null;
    if (!nextMethod) {
      return;
    }
    try {
      const payload = await batchUpdatePreferredMethod(selectedAccountIds, nextMethod);
      window.alert(`已切换 ${payload.summary.succeeded} 个档案。`);
      await loadAccounts();
    } catch (requestError) {
      window.alert(requestError instanceof Error ? requestError.message : '批量切换失败');
    }
  };

  return (
    <div className="relative flex h-full flex-1 flex-col overflow-hidden bg-white p-6 dark:bg-slate-950">
      <div className="mb-6 flex shrink-0 items-center justify-between">
        <div>
          <h1 className="text-2xl font-semibold text-slate-900 dark:text-slate-50">{t('accounts')}</h1>
          <p className="mt-1 text-sm text-slate-500">
            {t('total')}: {accounts.length}
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

      <div className="mb-4 flex shrink-0 items-center justify-between">
        <div className="relative w-72">
          <Search className="absolute left-3 top-2.5 h-4 w-4 text-slate-500" />
          <Input
            placeholder={t('search')}
            className="pl-9"
            value={searchQuery}
            onChange={(event) => setSearchQuery(event.target.value)}
          />
        </div>

        {selectedAccountIds.length > 0 ? (
          <div className="flex items-center gap-2 rounded-md bg-slate-100 px-3 py-1.5 dark:bg-slate-800">
            <span className="mr-2 text-sm font-medium">
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
            <Button variant="secondary" size="sm" onClick={() => void handleBatchChangeMethod()}>
              <Settings2 className="mr-2 h-4 w-4" />
              {t('batchChangeMethod')}
            </Button>
            <Button variant="destructive" size="sm" onClick={() => void handleBatchDelete()}>
              <Trash2 className="mr-2 h-4 w-4" />
              {t('batchDelete')}
            </Button>
          </div>
        ) : null}
      </div>

      {error ? <div className="mb-4 text-sm text-red-600 dark:text-red-400">{error}</div> : null}

      <div className="flex flex-1 flex-col overflow-hidden rounded-lg border border-slate-200 dark:border-slate-800">
        <div className="flex-1 overflow-auto">
          <table className="w-full text-left text-sm">
            <thead className="border-b border-slate-200 bg-slate-50 text-xs uppercase text-slate-500 dark:border-slate-800 dark:bg-slate-900/50">
              <tr>
                <th className="w-12 p-4">
                  <input
                    type="checkbox"
                    className="rounded border-slate-300"
                    checked={filteredAccounts.length > 0 && selectedAccountIds.length === filteredAccounts.length}
                    onChange={(event) =>
                      setSelectedAccountIds(event.target.checked ? filteredAccounts.map((account) => account.id) : [])
                    }
                  />
                </th>
                <th className="p-4 font-medium">{t('email')}</th>
                <th className="p-4 font-medium">{t('method')}</th>
                <th className="p-4 font-medium">{t('status')}</th>
                <th className="p-4 font-medium">{t('remark')}</th>
                <th className="p-4 text-right font-medium">{t('actions')}</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 dark:divide-slate-800">
              {loading ? (
                <tr>
                  <td colSpan={6} className="p-8 text-center text-slate-500">
                    {t('loadingAccounts')}
                  </td>
                </tr>
              ) : filteredAccounts.length === 0 ? (
                <tr>
                  <td colSpan={6} className="p-8 text-center text-slate-500">
                    {t('noAccounts')}
                  </td>
                </tr>
              ) : (
                filteredAccounts.map((account) => (
                  <tr key={account.id} className="transition-colors hover:bg-slate-50 dark:hover:bg-slate-900/50">
                    <td className="p-4">
                      <input
                        type="checkbox"
                        className="rounded border-slate-300"
                        checked={selectedAccountIds.includes(account.id)}
                        onChange={() => handleToggleAccount(account.id)}
                      />
                    </td>
                    <td className="p-4">
                      <div className="font-medium text-slate-900 dark:text-slate-100">{account.email}</div>
                      <div className="text-xs text-slate-500">{account.label}</div>
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
                            className={account.status === 'connected' ? 'h-4 w-4 text-emerald-500' : 'h-4 w-4 text-slate-400'}
                          />
                        )}
                        <span
                          className={
                            account.status === 'connected'
                              ? 'text-emerald-600 dark:text-emerald-400'
                              : account.status === 'disconnected'
                                ? 'text-red-600 dark:text-red-400'
                                : 'text-slate-500'
                          }
                        >
                          {t(account.status)}
                        </span>
                      </div>
                    </td>
                    <td className="p-4 text-slate-500">{account.notes || '-'}</td>
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
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/50">
          <div className="w-full max-w-xl rounded-xl border border-slate-200 bg-white p-6 shadow-lg dark:border-slate-800 dark:bg-slate-900">
            <h2 className="mb-4 text-lg font-semibold">
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
                    className="flex h-9 w-full rounded-md border border-slate-200 bg-transparent px-3 py-1 text-sm shadow-sm transition-colors dark:border-slate-800"
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
                  className="min-h-24 w-full rounded-md border border-slate-200 bg-transparent px-3 py-2 text-sm shadow-sm transition-colors focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-slate-950 dark:border-slate-800 dark:focus-visible:ring-slate-300"
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
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/50 px-4 py-6">
          <div className="flex max-h-full w-full max-w-2xl flex-col overflow-hidden rounded-xl border border-slate-200 bg-white shadow-lg dark:border-slate-800 dark:bg-slate-900">
            <div className="border-b border-slate-200 px-6 py-4 dark:border-slate-800">
              <h2 className="text-lg font-semibold">{t('importAccounts')}</h2>
              <p className="mt-1 text-sm text-slate-500">{t('importLineHint')}</p>
            </div>
            <form onSubmit={handleImport} className="flex min-h-0 flex-1 flex-col">
              <div className="min-h-0 space-y-4 overflow-y-auto px-6 py-5">
                <div>
                  <label className="mb-1 block text-sm font-medium">{t('defaultMethod')}</label>
                  <select
                    className="flex h-9 w-full rounded-md border border-slate-200 bg-transparent px-3 py-1 text-sm shadow-sm transition-colors dark:border-slate-800"
                    value={importForm.preferredMethod}
                    onChange={(event) =>
                      setImportForm((current) => ({
                        ...current,
                        preferredMethod: event.target.value as MethodValue,
                      }))
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
                  <label className="mb-1 block text-sm font-medium">{t('importPayload')}</label>
                  <textarea
                    className="min-h-64 w-full rounded-md border border-slate-200 bg-transparent px-3 py-2 font-mono text-sm shadow-sm transition-colors focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-slate-950 dark:border-slate-800 dark:focus-visible:ring-slate-300"
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

                <div className="rounded-lg border border-dashed border-slate-200 bg-slate-50 px-4 py-3 text-xs leading-6 text-slate-500 dark:border-slate-800 dark:bg-slate-950/40">
                  <div>{t('importSupportedFormats')}</div>
                  <div>{t('importExampleSimple')}</div>
                  <div>{t('importExampleKeyed')}</div>
                </div>

                {importError ? <div className="text-sm text-red-600 dark:text-red-400">{importError}</div> : null}
              </div>

              <div className="flex justify-end gap-2 border-t border-slate-200 px-6 py-4 dark:border-slate-800">
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
    </div>
  );
}
