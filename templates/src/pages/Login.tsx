import React, { useState } from 'react';
import { Navigate, useNavigate } from 'react-router-dom';
import { useAppStore } from '../store/useAppStore';
import { useI18n } from '../i18n';
import { Button } from '../components/ui/Button';
import { Input } from '../components/ui/Input';
import { BrandLockup, GitHubLink } from '../components/Brand';
import { APP_NAME } from '../lib/brand';
import { login as loginRequest } from '../lib/api';

export default function Login() {
  const { login, isAdmin, authReady } = useAppStore();
  const { t } = useI18n();
  const navigate = useNavigate();
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState('');

  if (authReady && isAdmin) {
    return <Navigate to="/" replace />;
  }

  const handleLogin = async (event: React.FormEvent) => {
    event.preventDefault();
    setSubmitting(true);
    setError('');
    try {
      const payload = await loginRequest({ username, password });
      login(payload.username);
      navigate('/');
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : t('loginFailed'));
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="relative flex min-h-[100dvh] items-center justify-center bg-black px-4">
      <div className="pointer-events-none absolute inset-0 bg-[radial-gradient(ellipse_at_top,rgba(255,255,255,0.08),transparent_55%)]" />
      <div className="relative w-full max-w-[380px]">
        <div className="mb-8 flex flex-col items-center gap-4 text-center">
          <BrandLockup />
          <p className="text-sm text-white/50">{t('adminLogin')}</p>
        </div>
        <div className="rounded-xl border border-white/10 bg-[#0a0a0a] p-1.5">
          <form
            onSubmit={handleLogin}
            className="space-y-4 rounded-[10px] border border-white/10 bg-black px-5 py-6"
            autoComplete="on"
          >
            <div className="space-y-2">
              <label className="text-xs text-white/50" htmlFor="admin-username">
                {t('username')}
              </label>
              <Input
                id="admin-username"
                type="text"
                name="username"
                autoComplete="username"
                autoCapitalize="none"
                autoFocus
                spellCheck={false}
                value={username}
                onChange={(event) => setUsername(event.target.value)}
                required
              />
            </div>
            <div className="space-y-2">
              <label className="text-xs text-white/50" htmlFor="admin-password">
                {t('password')}
              </label>
              <div className="relative">
                <Input
                  id="admin-password"
                  type={showPassword ? 'text' : 'password'}
                  name="password"
                  autoComplete="current-password"
                  value={password}
                  onChange={(event) => setPassword(event.target.value)}
                  required
                  className="pr-16"
                />
                <button
                  type="button"
                  className="absolute inset-y-0 right-2 text-xs text-white/45 transition duration-200 ease-[cubic-bezier(0.32,0.72,0,1)] hover:text-white"
                  onClick={() => setShowPassword((current) => !current)}
                >
                  {showPassword ? '隐藏' : '显示'}
                </button>
              </div>
            </div>
            {error ? <div className="text-sm text-[#ff8080]">{error}</div> : null}
            <Button type="submit" className="w-full" disabled={submitting}>
              {submitting ? t('loggingIn') : t('login')}
            </Button>
          </form>
        </div>
        <div className="mt-6 flex items-center justify-center">
          <GitHubLink />
        </div>
        <p className="mt-3 text-center text-[11px] text-white/30">{APP_NAME}</p>
      </div>
    </div>
  );
}
