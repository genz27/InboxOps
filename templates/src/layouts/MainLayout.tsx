import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Navigate, Outlet, useLocation, useNavigate } from 'react-router-dom';
import { Check, ChevronDown, Copy, Languages, LogOut, Settings2 } from 'lucide-react';
import { useAppStore } from '../store/useAppStore';
import { useI18n } from '../i18n';
import { Button } from '../components/ui/Button';
import { Input } from '../components/ui/Input';
import { BrandLockup, GitHubMark } from '../components/Brand';
import { GITHUB_REPO_LABEL, GITHUB_REPO_URL } from '../lib/brand';
import { useFeedback } from '../components/Feedback';
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
  const navigate = useNavigate();
  const location = useLocation();
  const { toast } = useFeedback();
  const [userMenuOpen, setUserMenuOpen] = useState(false);
  const userMenuRef = useRef<HTMLDivElement | null>(null);
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

  useEffect(() => {
    setCopyState('idle');
  }, [activeMailboxEmail]);

  useEffect(() => {
    if (!userMenuOpen) {
      return;
    }
    const onPointerDown = (event: MouseEvent) => {
      if (!userMenuRef.current?.contains(event.target as Node)) {
        setUserMenuOpen(false);
      }
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setUserMenuOpen(false);
      }
    };
    window.addEventListener('mousedown', onPointerDown);
    window.addEventListener('keydown', onKeyDown);
    return () => {
      window.removeEventListener('mousedown', onPointerDown);
      window.removeEventListener('keydown', onKeyDown);
    };
  }, [userMenuOpen]);

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
      toast(t('passwordChangeSuccess'));
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
      toast(t('copyMailboxEmailFailed'), 'error');
    }
  };

  return (
    <div className="relative flex h-screen w-full flex-col bg-black text-white">
      <header className="flex h-14 shrink-0 items-center justify-between border-b border-white/10 px-4">
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
        <div className="flex min-w-0 items-center gap-1">
          {activeMailboxEmail ? (
            <button
              type="button"
              className="hidden min-w-0 max-w-[14rem] items-center gap-1.5 rounded-md px-2 py-1 text-xs text-white/70 transition duration-200 ease-[cubic-bezier(0.32,0.72,0,1)] hover:bg-white/5 hover:text-white md:inline-flex"
              onClick={() => void handleCopyMailboxEmail()}
              title={copyState === 'success' ? t('copyMailboxEmailSuccess') : t('copyMailboxEmail')}
            >
              <span className="truncate">{activeMailboxEmail}</span>
              {copyState === 'success' ? <Check className="h-3 w-3 shrink-0" /> : <Copy className="h-3 w-3 shrink-0 opacity-50" />}
            </button>
          ) : null}
          <div className="relative" ref={userMenuRef}>
            <button
              type="button"
              className="inline-flex items-center gap-1 rounded-md px-2 py-1 text-sm text-white/80 transition duration-200 ease-[cubic-bezier(0.32,0.72,0,1)] hover:bg-white/5 hover:text-white"
              onClick={() => setUserMenuOpen((open) => !open)}
              aria-expanded={userMenuOpen}
            >
              <span className="max-w-[8rem] truncate">{username ?? 'admin'}</span>
              <ChevronDown className="h-3.5 w-3.5 text-white/40" />
            </button>
            {userMenuOpen ? (
              <div className="absolute right-0 top-full z-50 mt-1 w-52 overflow-hidden rounded-lg border border-white/10 bg-[#0a0a0a] py-1">
                <button
                  type="button"
                  className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-white/80 hover:bg-white/5"
                  onClick={() => {
                    setUserMenuOpen(false);
                    openPasswordDialog();
                  }}
                >
                  <Settings2 className="h-3.5 w-3.5" />
                  {t('changePassword')}
                </button>
                <button
                  type="button"
                  className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-white/80 hover:bg-white/5"
                  onClick={() => {
                    setLanguage(language === 'en' ? 'zh' : 'en');
                    setUserMenuOpen(false);
                  }}
                >
                  <Languages className="h-3.5 w-3.5" />
                  {language === 'en' ? t('chinese') : t('english')}
                </button>
                <a
                  href={GITHUB_REPO_URL}
                  target="_blank"
                  rel="noreferrer noopener"
                  className="flex w-full items-center gap-2 px-3 py-2 text-sm text-white/80 hover:bg-white/5"
                  onClick={() => setUserMenuOpen(false)}
                >
                  <GitHubMark className="h-3.5 w-3.5" />
                  {GITHUB_REPO_LABEL}
                </a>
                <button
                  type="button"
                  className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-white/80 hover:bg-white/5"
                  onClick={() => {
                    setUserMenuOpen(false);
                    void handleLogout();
                  }}
                >
                  <LogOut className="h-3.5 w-3.5" />
                  {t('logout')}
                </button>
              </div>
            ) : null}
          </div>
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
