import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Mail } from 'lucide-react';
import { useAppStore } from '../store/useAppStore';
import { useI18n } from '../i18n';
import { Button } from '../components/ui/Button';
import { Input } from '../components/ui/Input';
import { login as loginRequest } from '../lib/api';

export default function Login() {
  const { login } = useAppStore();
  const { t } = useI18n();
  const navigate = useNavigate();
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState('');

  const handleLogin = async (event: React.FormEvent) => {
    event.preventDefault();
    setSubmitting(true);
    setError('');
    try {
      const payload = await loginRequest({ username, password });
      login(payload.username);
      navigate('/');
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : '登录失败');
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="flex min-h-screen items-center justify-center bg-slate-50 dark:bg-slate-950">
      <div className="w-full max-w-md space-y-8 rounded-xl border border-slate-200 bg-white p-8 shadow-sm dark:border-slate-800 dark:bg-slate-900">
        <div className="flex flex-col items-center justify-center space-y-2 text-center">
          <div className="rounded-full bg-slate-100 p-3 dark:bg-slate-800">
            <Mail className="h-6 w-6 text-slate-900 dark:text-slate-50" />
          </div>
          <h1 className="text-2xl font-semibold tracking-tight text-slate-900 dark:text-slate-50">
            {t('adminLogin')}
          </h1>
        </div>
        <form onSubmit={handleLogin} className="space-y-4">
          <div className="space-y-2">
            <Input
              type="text"
              placeholder={t('username')}
              value={username}
              onChange={(event) => setUsername(event.target.value)}
              required
            />
          </div>
          <div className="space-y-2">
            <Input
              type="password"
              placeholder={t('password')}
              value={password}
              onChange={(event) => setPassword(event.target.value)}
              required
            />
          </div>
          {error ? <div className="text-sm text-red-600 dark:text-red-400">{error}</div> : null}
          <Button type="submit" className="w-full" disabled={submitting}>
            {submitting ? '登录中...' : t('login')}
          </Button>
        </form>
      </div>
    </div>
  );
}
