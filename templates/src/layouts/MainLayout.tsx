import React, { useEffect, useMemo, useState } from 'react';
import { Navigate, Outlet, useLocation, useNavigate } from 'react-router-dom';
import { Check, Copy, Languages, LogOut, Mail, Monitor, Moon, Settings2, Sun } from 'lucide-react';
import { useAppStore } from '../store/useAppStore';
import { useI18n } from '../i18n';
import { useTheme } from '../theme/useTheme';
import { Button } from '../components/ui/Button';
import { Input } from '../components/ui/Input';
import { changeAdminPassword, logout as logoutRequest } from '../lib/api';

interface PasswordFormState {
  currentPassword: string;
  newPassword: string;
  confirmPassword: string;
}

const DEFAULT_PASSWORD_FORM: PasswordFormState = {
  currentPassword: '',
  newPassword: '',
  confirmPassword: '',
};

async function copyToClipboard(text: string) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.setAttribute('readonly', 'true');
  textarea.style.position = 'fixed';
  textarea.style.opacity = '0';
  document.body.appendChild(textarea);
  textarea.focus();
  textarea.select();

  const copied = document.execCommand('copy');
  textarea.remove();

  if (!copied) {
    throw new Error('copy_failed');
  }
}

export default function MainLayout() {
  const { isAdmin, authReady, logout, username, accounts, activeMailboxId } = useAppStore();
  const { t, language, setLanguage } = useI18n();
  const { theme, setTheme } = useTheme();
  const navigate = useNavigate();
  const location = useLocation();
  const [passwordDialogOpen, setPasswordDialogOpen] = useState(false);
  const [passwordForm, setPasswordForm] = useState<PasswordFormState>(DEFAULT_PASSWORD_FORM);
  const [passwordSaving, setPasswordSaving] = useState(false);
  const [passwordError, setPasswordError] = useState('');
  const [copyState, setCopyState] = useState<'idle' | 'success'>('idle');

  const activeMailboxEmail = useMemo(
    () => accounts.find((account) => account.id === activeMailboxId)?.email ?? '',
    [accounts, activeMailboxId],
  );

  useEffect(() => {
    setCopyState('idle');
  }, [activeMailboxEmail]);

  useEffect(() => {
    if (copyState !== 'success') {
      return;
    }

    const timerId = window.setTimeout(() => {
      setCopyState('idle');
    }, 1500);

    return () => {
      window.clearTimeout(timerId);
    };
  }, [copyState]);

  if (!authReady) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-slate-50 text-sm text-slate-500 dark:bg-slate-950 dark:text-slate-400">
        正在校验登录状态...
      </div>
    );
  }

  if (!isAdmin) {
    return <Navigate to="/login" replace />;
  }

  const handleLogout = async () => {
    try {
      await logoutRequest();
    } finally {
      logout();
      navigate('/login');
    }
  };

  const closePasswordDialog = () => {
    setPasswordDialogOpen(false);
    setPasswordForm(DEFAULT_PASSWORD_FORM);
    setPasswordError('');
  };

  const openPasswordDialog = () => {
    setPasswordDialogOpen(true);
    setPasswordForm(DEFAULT_PASSWORD_FORM);
    setPasswordError('');
  };

  const handlePasswordChange = async (event: React.FormEvent) => {
    event.preventDefault();
    setPasswordError('');

    if (passwordForm.newPassword.length < 8) {
      setPasswordError(t('passwordTooShort'));
      return;
    }
    if (passwordForm.newPassword !== passwordForm.confirmPassword) {
      setPasswordError(t('passwordMismatch'));
      return;
    }

    setPasswordSaving(true);
    try {
      await changeAdminPassword(passwordForm);
      closePasswordDialog();
      window.alert(t('passwordChangeSuccess'));
    } catch (requestError) {
      setPasswordError(requestError instanceof Error ? requestError.message : t('passwordChangeFailed'));
    } finally {
      setPasswordSaving(false);
    }
  };

  const handleCopyMailboxEmail = async () => {
    if (!activeMailboxEmail) {
      return;
    }

    try {
      await copyToClipboard(activeMailboxEmail);
      setCopyState('success');
    } catch {
      window.alert(t('copyMailboxEmailFailed'));
    }
  };

  return (
    <div className="relative flex h-screen w-full flex-col bg-slate-50 text-slate-900 dark:bg-slate-950 dark:text-slate-50">
      <header className="flex h-14 shrink-0 items-center justify-between border-b border-slate-200 px-4 dark:border-slate-800">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2 font-semibold">
            <Mail className="h-5 w-5" />
            <span>Email Outlook</span>
          </div>
          <nav className="ml-6 flex items-center gap-2">
            <Button
              variant={location.pathname === '/' ? 'secondary' : 'ghost'}
              size="sm"
              onClick={() => navigate('/')}
            >
              {t('workspace')}
            </Button>
            <Button
              variant={location.pathname === '/accounts' ? 'secondary' : 'ghost'}
              size="sm"
              onClick={() => navigate('/accounts')}
            >
              {t('accounts')}
            </Button>
          </nav>
        </div>
        <div className="flex items-center gap-2">
          <div className="hidden text-xs text-slate-500 dark:text-slate-400 md:block">{username ?? 'admin'}</div>
          <div className="hidden max-w-[28rem] items-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-1.5 text-xs dark:border-slate-800 dark:bg-slate-900 lg:flex">
            <span className="shrink-0 text-slate-500 dark:text-slate-400">{t('currentMailbox')}</span>
            <span className="min-w-0 flex-1 break-all font-medium text-slate-800 dark:text-slate-100">
              {activeMailboxEmail || t('noMailboxSelected')}
            </span>
            <button
              type="button"
              className="inline-flex h-6 w-6 shrink-0 items-center justify-center rounded-md text-slate-500 transition-colors hover:bg-slate-100 hover:text-slate-900 disabled:cursor-not-allowed disabled:opacity-40 dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-slate-100"
              onClick={() => void handleCopyMailboxEmail()}
              title={copyState === 'success' ? t('copyMailboxEmailSuccess') : t('copyMailboxEmail')}
              aria-label={copyState === 'success' ? t('copyMailboxEmailSuccess') : t('copyMailboxEmail')}
              disabled={!activeMailboxEmail}
            >
              {copyState === 'success' ? <Check className="h-3.5 w-3.5 text-emerald-500" /> : <Copy className="h-3.5 w-3.5" />}
            </button>
          </div>
          <Button variant="ghost" size="sm" onClick={openPasswordDialog} title={t('changePassword')}>
            <Settings2 className="mr-2 h-4 w-4" />
            {t('changePassword')}
          </Button>
          <Button
            variant="ghost"
            size="icon"
            onClick={() => setLanguage(language === 'en' ? 'zh' : 'en')}
            title={t('language')}
          >
            <Languages className="h-4 w-4" />
          </Button>
          <Button
            variant="ghost"
            size="icon"
            onClick={() => {
              const nextTheme = theme === 'light' ? 'dark' : theme === 'dark' ? 'system' : 'light';
              setTheme(nextTheme);
            }}
            title={t('theme')}
          >
            {theme === 'light' ? (
              <Sun className="h-4 w-4" />
            ) : theme === 'dark' ? (
              <Moon className="h-4 w-4" />
            ) : (
              <Monitor className="h-4 w-4" />
            )}
          </Button>
          <Button variant="ghost" size="sm" onClick={handleLogout}>
            <LogOut className="mr-2 h-4 w-4" />
            {t('logout')}
          </Button>
        </div>
      </header>
      <main className="flex-1 overflow-hidden">
        <Outlet />
      </main>

      {passwordDialogOpen ? (
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/50 px-4 py-6">
          <div className="w-full max-w-md rounded-xl border border-slate-200 bg-white p-6 shadow-lg dark:border-slate-800 dark:bg-slate-900">
            <h2 className="text-lg font-semibold">{t('changePassword')}</h2>
            <p className="mt-1 text-sm text-slate-500">{t('passwordChangeHint')}</p>
            <div className="mt-4 flex items-start gap-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-sm dark:border-slate-800 dark:bg-slate-950/60">
              <div className="min-w-0 flex-1">
                <div className="text-xs text-slate-500 dark:text-slate-400">{t('currentMailbox')}</div>
                <div className="mt-1 break-all font-medium text-slate-900 dark:text-slate-100">
                  {activeMailboxEmail || t('noMailboxSelected')}
                </div>
              </div>
              <button
                type="button"
                className="inline-flex h-8 w-8 shrink-0 items-center justify-center rounded-md text-slate-500 transition-colors hover:bg-slate-200 hover:text-slate-900 disabled:cursor-not-allowed disabled:opacity-40 dark:text-slate-400 dark:hover:bg-slate-800 dark:hover:text-slate-100"
                onClick={() => void handleCopyMailboxEmail()}
                title={copyState === 'success' ? t('copyMailboxEmailSuccess') : t('copyMailboxEmail')}
                aria-label={copyState === 'success' ? t('copyMailboxEmailSuccess') : t('copyMailboxEmail')}
                disabled={!activeMailboxEmail}
              >
                {copyState === 'success' ? <Check className="h-4 w-4 text-emerald-500" /> : <Copy className="h-4 w-4" />}
              </button>
            </div>

            <form onSubmit={handlePasswordChange} className="mt-5 space-y-4">
              <div>
                <label className="mb-1 block text-sm font-medium">{t('currentPassword')}</label>
                <Input
                  type="password"
                  value={passwordForm.currentPassword}
                  onChange={(event) =>
                    setPasswordForm((current) => ({
                      ...current,
                      currentPassword: event.target.value,
                    }))
                  }
                  autoComplete="current-password"
                  required
                />
              </div>
              <div>
                <label className="mb-1 block text-sm font-medium">{t('newPassword')}</label>
                <Input
                  type="password"
                  value={passwordForm.newPassword}
                  onChange={(event) =>
                    setPasswordForm((current) => ({
                      ...current,
                      newPassword: event.target.value,
                    }))
                  }
                  autoComplete="new-password"
                  required
                />
              </div>
              <div>
                <label className="mb-1 block text-sm font-medium">{t('confirmPassword')}</label>
                <Input
                  type="password"
                  value={passwordForm.confirmPassword}
                  onChange={(event) =>
                    setPasswordForm((current) => ({
                      ...current,
                      confirmPassword: event.target.value,
                    }))
                  }
                  autoComplete="new-password"
                  required
                />
              </div>

              {passwordError ? <div className="text-sm text-red-600 dark:text-red-400">{passwordError}</div> : null}

              <div className="flex justify-end gap-2">
                <Button type="button" variant="outline" onClick={closePasswordDialog} disabled={passwordSaving}>
                  {t('cancel')}
                </Button>
                <Button type="submit" disabled={passwordSaving}>
                  {passwordSaving ? t('savingPassword') : t('save')}
                </Button>
              </div>
            </form>
          </div>
        </div>
      ) : null}
    </div>
  );
}
