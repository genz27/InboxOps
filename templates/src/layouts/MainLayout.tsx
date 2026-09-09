import React, { useEffect, useMemo, useState } from 'react';
import { Navigate, Outlet, useLocation, useNavigate } from 'react-router-dom';
import { Check, Copy, Languages, LogOut, Monitor, Moon, Settings2, Sun } from 'lucide-react';
import { useAppStore } from '../store/useAppStore';
import { useI18n } from '../i18n';
import { useTheme } from '../theme/useTheme';
import { Button } from '../components/ui/Button';
import { Input } from '../components/ui/Input';
import { BrandLockup, GitHubLink } from '../components/Brand';
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
  const { isAdmin, authReady, logout, username, accounts, activeMailboxId, folders, activeFolderId } = useAppStore();
  const { t, language, setLanguage } = useI18n();
  const { theme, setTheme } = useTheme();
  const navigate = useNavigate();
  const location = useLocation();
  const [passwordDialogOpen, setPasswordDialogOpen] = useState(false);
  const [passwordForm, setPasswordForm] = useState<PasswordFormState>(DEFAULT_PASSWORD_FORM);
  const [passwordSaving, setPasswordSaving] = useState(false);
  const [passwordError, setPasswordError] = useState('');
  const [copyState, setCopyState] = useState<'idle' | 'success'>('idle');

  const activeMailbox = useMemo(
    () => accounts.find((account) => account.id === activeMailboxId) ?? null,
    [accounts, activeMailboxId],
  );
  const activeMailboxEmail = activeMailbox?.email ?? '';
  const activeFolderName = useMemo(
    () => folders.find((folder) => folder.id === activeFolderId)?.displayName ?? '',
    [folders, activeFolderId],
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
      <div className="flex min-h-[100dvh] items-center justify-center bg-black text-sm text-white/45">
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
    <div className="relative flex h-screen w-full flex-col bg-black text-white">
      <header className="flex h-12 shrink-0 items-center justify-between border-b border-white/10 px-4">
        <div className="flex items-center gap-4">
          <BrandLockup compact />
          <nav className="ml-2 flex items-center gap-1">
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
        <div className="flex items-center gap-1.5">
          <GitHubLink className="hidden md:inline-flex" />
          <div className="hidden text-xs text-white/40 md:block">{username ?? 'admin'}</div>
          <div className="hidden max-w-[28rem] items-center gap-2 rounded-md border border-white/10 bg-[#0a0a0a] px-3 py-1 text-xs lg:flex">
            <span className="shrink-0 text-white/40">{t('currentMailbox')}</span>
            <span className="min-w-0 flex-1 truncate font-medium text-white/90" title={activeMailboxEmail}>
              {activeMailboxEmail || t('noMailboxSelected')}
            </span>
            {activeFolderName ? (
              <span className="hidden shrink-0 rounded bg-white/10 px-1.5 py-0.5 text-[10px] text-white/60 xl:inline">
                {activeFolderName}
              </span>
            ) : null}
            {activeMailbox?.preferredMethod || activeMailbox?.method ? (
              <span className="hidden shrink-0 rounded bg-white/10 px-1.5 py-0.5 text-[10px] text-white/60 xl:inline">
                {activeMailbox.preferredMethod || activeMailbox.method}
              </span>
            ) : null}
            <button
              type="button"
              className="inline-flex h-6 w-6 shrink-0 items-center justify-center rounded-md text-white/45 transition duration-200 ease-[cubic-bezier(0.32,0.72,0,1)] hover:bg-white/10 hover:text-white disabled:cursor-not-allowed disabled:opacity-40"
              onClick={() => void handleCopyMailboxEmail()}
              title={copyState === 'success' ? t('copyMailboxEmailSuccess') : t('copyMailboxEmail')}
              aria-label={copyState === 'success' ? t('copyMailboxEmailSuccess') : t('copyMailboxEmail')}
              disabled={!activeMailboxEmail}
            >
              {copyState === 'success' ? <Check className="h-3.5 w-3.5 text-white" /> : <Copy className="h-3.5 w-3.5" />}
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
        <div className="absolute inset-0 z-50 flex items-center justify-center bg-black/70 px-4 py-6">
          <div className="w-full max-w-md rounded-xl border border-white/10 bg-[#0a0a0a] p-6">
            <h2 className="text-lg font-medium tracking-tight">{t('changePassword')}</h2>
            <p className="mt-1 text-sm text-white/45">{t('passwordChangeHint')}</p>
            <div className="mt-4 flex items-start gap-2 rounded-lg border border-white/10 bg-black px-3 py-2 text-sm">
              <div className="min-w-0 flex-1">
                <div className="text-xs text-white/40">{t('currentMailbox')}</div>
                <div className="mt-1 break-all font-medium text-white">
                  {activeMailboxEmail || t('noMailboxSelected')}
                </div>
              </div>
              <button
                type="button"
                className="inline-flex h-8 w-8 shrink-0 items-center justify-center rounded-md text-white/45 transition duration-200 ease-[cubic-bezier(0.32,0.72,0,1)] hover:bg-white/10 hover:text-white disabled:cursor-not-allowed disabled:opacity-40"
                onClick={() => void handleCopyMailboxEmail()}
                title={copyState === 'success' ? t('copyMailboxEmailSuccess') : t('copyMailboxEmail')}
                aria-label={copyState === 'success' ? t('copyMailboxEmailSuccess') : t('copyMailboxEmail')}
                disabled={!activeMailboxEmail}
              >
                {copyState === 'success' ? <Check className="h-4 w-4 text-white" /> : <Copy className="h-4 w-4" />}
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

              {passwordError ? <div className="text-sm text-[#ff8080]">{passwordError}</div> : null}

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
