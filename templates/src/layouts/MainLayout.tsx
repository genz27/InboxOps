import React from 'react';
import { Navigate, Outlet, useLocation, useNavigate } from 'react-router-dom';
import { Languages, LogOut, Mail, Monitor, Moon, Sun } from 'lucide-react';
import { useAppStore } from '../store/useAppStore';
import { useI18n } from '../i18n';
import { useTheme } from '../theme/useTheme';
import { Button } from '../components/ui/Button';
import { logout as logoutRequest } from '../lib/api';

export default function MainLayout() {
  const { isAdmin, authReady, logout, username } = useAppStore();
  const { t, language, setLanguage } = useI18n();
  const { theme, setTheme } = useTheme();
  const navigate = useNavigate();
  const location = useLocation();

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

  return (
    <div className="flex h-screen w-full flex-col bg-slate-50 text-slate-900 dark:bg-slate-950 dark:text-slate-50">
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
    </div>
  );
}
