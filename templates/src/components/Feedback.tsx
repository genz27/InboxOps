import React, { createContext, useCallback, useContext, useMemo, useRef, useState } from 'react';
import { Button } from './ui/Button';
import { Input } from './ui/Input';
import { cn } from '../lib/utils';

type ToastKind = 'ok' | 'error';

interface ToastItem {
  id: string;
  kind: ToastKind;
  message: string;
}

interface ConfirmOptions {
  title: string;
  body: string;
  confirmLabel?: string;
  danger?: boolean;
}

interface PromptOptions {
  title: string;
  body?: string;
  label?: string;
  defaultValue?: string;
  confirmLabel?: string;
}

interface FeedbackContextValue {
  toast: (message: string, kind?: ToastKind) => void;
  confirm: (options: ConfirmOptions) => Promise<boolean>;
  prompt: (options: PromptOptions) => Promise<string | null>;
}

const FeedbackContext = createContext<FeedbackContextValue | null>(null);

export function useFeedback(): FeedbackContextValue {
  const value = useContext(FeedbackContext);
  if (!value) {
    throw new Error('useFeedback must be used within FeedbackProvider');
  }
  return value;
}

export function FeedbackProvider({ children }: { children: React.ReactNode }) {
  const [toasts, setToasts] = useState<ToastItem[]>([]);
  const [confirmState, setConfirmState] = useState<(ConfirmOptions & { resolve: (value: boolean) => void }) | null>(
    null,
  );
  const [promptState, setPromptState] = useState<
    (PromptOptions & { resolve: (value: string | null) => void }) | null
  >(null);
  const [promptValue, setPromptValue] = useState('');
  const promptRef = useRef<HTMLInputElement | null>(null);
  const timers = useRef<number[]>([]);

  const toast = useCallback((message: string, kind: ToastKind = 'ok') => {
    const id = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    setToasts((current) => [...current.slice(-2), { id, kind, message }]);
    const timerId = window.setTimeout(() => {
      setToasts((current) => current.filter((item) => item.id !== id));
    }, 3200);
    timers.current.push(timerId);
  }, []);

  const confirm = useCallback((options: ConfirmOptions) => {
    return new Promise<boolean>((resolve) => {
      setConfirmState({ ...options, resolve });
    });
  }, []);

  const prompt = useCallback((options: PromptOptions) => {
    return new Promise<string | null>((resolve) => {
      setPromptValue(options.defaultValue ?? '');
      setPromptState({ ...options, resolve });
      window.setTimeout(() => promptRef.current?.focus(), 30);
    });
  }, []);

  const value = useMemo(() => ({ toast, confirm, prompt }), [toast, confirm, prompt]);

  return (
    <FeedbackContext.Provider value={value}>
      {children}
      <div className="pointer-events-none fixed bottom-5 left-1/2 z-[60] flex w-[min(92vw,28rem)] -translate-x-1/2 flex-col gap-2">
        {toasts.map((item) => (
          <div
            key={item.id}
            className={cn(
              'pointer-events-auto rounded-lg border px-3 py-2 text-sm shadow-[0_8px_30px_rgba(0,0,0,0.35)]',
              item.kind === 'error'
                ? 'border-[#ff4d4d]/30 bg-[#1a0a0a] text-[#ff8080]'
                : 'border-white/10 bg-[#111] text-white',
            )}
          >
            {item.message}
          </div>
        ))}
      </div>
      {confirmState ? (
        <div className="fixed inset-0 z-[70] flex items-center justify-center bg-black/70 px-4">
          <div className="w-full max-w-sm rounded-xl border border-white/10 bg-[#0a0a0a] p-5">
            <h2 className="text-base font-medium tracking-tight text-white">{confirmState.title}</h2>
            <p className="mt-2 text-sm leading-6 text-white/55">{confirmState.body}</p>
            <div className="mt-5 flex justify-end gap-2">
              <Button
                type="button"
                variant="outline"
                onClick={() => {
                  confirmState.resolve(false);
                  setConfirmState(null);
                }}
              >
                取消
              </Button>
              <Button
                type="button"
                variant={confirmState.danger ? 'destructive' : 'default'}
                onClick={() => {
                  confirmState.resolve(true);
                  setConfirmState(null);
                }}
              >
                {confirmState.confirmLabel || '确认'}
              </Button>
            </div>
          </div>
        </div>
      ) : null}
      {promptState ? (
        <div className="fixed inset-0 z-[70] flex items-center justify-center bg-black/70 px-4">
          <form
            className="w-full max-w-sm rounded-xl border border-white/10 bg-[#0a0a0a] p-5"
            onSubmit={(event) => {
              event.preventDefault();
              const next = promptValue.trim();
              promptState.resolve(next ? next : null);
              setPromptState(null);
            }}
          >
            <h2 className="text-base font-medium tracking-tight text-white">{promptState.title}</h2>
            {promptState.body ? <p className="mt-2 text-sm text-white/55">{promptState.body}</p> : null}
            <label className="mt-4 block text-xs text-white/45">{promptState.label || '名称'}</label>
            <Input
              ref={promptRef}
              className="mt-1.5"
              value={promptValue}
              onChange={(event) => setPromptValue(event.target.value)}
            />
            <div className="mt-5 flex justify-end gap-2">
              <Button
                type="button"
                variant="outline"
                onClick={() => {
                  promptState.resolve(null);
                  setPromptState(null);
                }}
              >
                取消
              </Button>
              <Button type="submit">{promptState.confirmLabel || '确认'}</Button>
            </div>
          </form>
        </div>
      ) : null}
    </FeedbackContext.Provider>
  );
}
