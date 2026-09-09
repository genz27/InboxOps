import React from 'react';
import { APP_NAME, GITHUB_REPO_LABEL, GITHUB_REPO_URL } from '../lib/brand';
import { cn } from '../lib/utils';

export function BrandMark({ className }: { className?: string }) {
  return (
    <span
      className={cn(
        'inline-flex h-7 w-7 items-center justify-center rounded-md bg-white text-black',
        className,
      )}
      aria-hidden
    >
      <svg viewBox="0 0 24 24" className="h-3.5 w-3.5" fill="none">
        <path
          d="M3.5 7.5 12 3.5l8.5 4v9L12 20.5 3.5 16.5v-9Z"
          stroke="currentColor"
          strokeWidth="1.5"
        />
        <path d="M3.5 7.5 12 12l8.5-4.5" stroke="currentColor" strokeWidth="1.5" />
      </svg>
    </span>
  );
}

export function GitHubMark({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 16 16" className={cn('h-4 w-4', className)} fill="currentColor" aria-hidden>
      <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82A7.68 7.68 0 0 1 8 4.07c.68.003 1.36.092 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0 0 16 8c0-4.42-3.58-8-8-8Z" />
    </svg>
  );
}

export function BrandLockup({
  compact = false,
  className,
}: {
  compact?: boolean;
  className?: string;
}) {
  return (
    <div className={cn('flex items-center gap-2.5', className)}>
      <BrandMark />
      <div className="min-w-0 leading-tight">
        <div className={compact ? 'truncate text-sm font-medium tracking-tight text-white' : 'truncate text-lg font-medium tracking-tight text-white'}>
          {APP_NAME}
        </div>
      </div>
    </div>
  );
}

export function GitHubLink({ className }: { className?: string }) {
  return (
    <a
      href={GITHUB_REPO_URL}
      target="_blank"
      rel="noreferrer noopener"
      className={cn(
        'inline-flex items-center gap-1.5 rounded-md px-2 py-1 text-xs text-white/55 transition duration-300 ease-[cubic-bezier(0.32,0.72,0,1)] hover:bg-white/5 hover:text-white',
        className,
      )}
    >
      <GitHubMark />
      <span>{GITHUB_REPO_LABEL}</span>
    </a>
  );
}
