import React, { useState } from 'react';
import { Navigate, Outlet, useLocation, useNavigate } from 'react-router-dom';
import { Languages, LogOut, Mail, Monitor, Moon, Settings2, Sun } from 'lucide-react';
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

export default function MainLayout() {
  const { isAdmin, authReady, logout, username } = useAppStore();
  const { t, language, setLanguage } = useI18n();
  const { theme, setTheme } = useTheme();
  const navigate = useNavigate();
  const location = useLocation();
  const [passwordDialogOpen, setPasswordDialogOpen] = useState(false);
  const [passwordForm, setPasswordForm] = useState<PasswordFormState>(DEFAULT_PASSWORD_FORM);
  const [passwordSaving, setPasswordSaving] = useState(false);
  const [passwordError, setPasswordError] = useState('');

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
