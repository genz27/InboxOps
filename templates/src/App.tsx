import React, { useEffect } from 'react';
import { BrowserRouter, Navigate, Route, Routes } from 'react-router-dom';
import MainLayout from './layouts/MainLayout';
import Login from './pages/Login';
import Workspace from './pages/Workspace';
import Accounts from './pages/Accounts';
import { getAuthMe } from './lib/api';
import { useAppStore } from './store/useAppStore';

function AuthBootstrap() {
  const { authReady, setAuthState } = useAppStore();

  useEffect(() => {
    let cancelled = false;

    void (async () => {
      try {
        const payload = await getAuthMe();
        if (!cancelled) {
          setAuthState(payload.authenticated, payload.username);
        }
      } catch {
        if (!cancelled) {
          setAuthState(false, null);
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [setAuthState]);

  if (!authReady) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-slate-50 text-sm text-slate-500 dark:bg-slate-950 dark:text-slate-400">
        正在初始化前端...
      </div>
    );
  }

  return (
    <Routes>
      <Route path="/login" element={<Login />} />
      <Route element={<MainLayout />}>
        <Route path="/" element={<Workspace />} />
        <Route path="/accounts" element={<Accounts />} />
      </Route>
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  );
}

export default function App() {
  return (
    <BrowserRouter>
      <AuthBootstrap />
    </BrowserRouter>
  );
}
