const PINNED_KEY = 'inboxops.pinnedMailboxes';
const RECENT_KEY = 'inboxops.recentMailboxes';
const NOTIFY_KEY = 'inboxops.notifyOtp';

function readStringArray(key: string): string[] {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw) as unknown;
    if (!Array.isArray(parsed)) {
      return [];
    }
    return parsed.filter((item): item is string => typeof item === 'string' && item.length > 0);
  } catch {
    return [];
  }
}

function writeStringArray(key: string, values: string[]) {
  window.localStorage.setItem(key, JSON.stringify(values));
}

export function getPinnedMailboxIds(): string[] {
  return readStringArray(PINNED_KEY);
}

export function togglePinnedMailbox(mailboxId: string): string[] {
  const current = new Set(getPinnedMailboxIds());
  if (current.has(mailboxId)) {
    current.delete(mailboxId);
  } else {
    current.add(mailboxId);
  }
  const next = Array.from(current);
  writeStringArray(PINNED_KEY, next);
  return next;
}

export function getRecentMailboxIds(): string[] {
  return readStringArray(RECENT_KEY);
}

export function touchRecentMailbox(mailboxId: string, limit = 12): string[] {
  const next = [mailboxId, ...getRecentMailboxIds().filter((id) => id !== mailboxId)].slice(0, limit);
  writeStringArray(RECENT_KEY, next);
  return next;
}

export function sortMailboxesByPreference<T extends { id: string }>(
  accounts: T[],
  pinnedIds: string[],
  recentIds: string[],
): T[] {
  const pinned = new Set(pinnedIds);
  const recentRank = new Map(recentIds.map((id, index) => [id, index]));
  return [...accounts].sort((left, right) => {
    const leftPinned = pinned.has(left.id) ? 0 : 1;
    const rightPinned = pinned.has(right.id) ? 0 : 1;
    if (leftPinned !== rightPinned) {
      return leftPinned - rightPinned;
    }
    const leftRecent = recentRank.has(left.id) ? recentRank.get(left.id)! : Number.MAX_SAFE_INTEGER;
    const rightRecent = recentRank.has(right.id) ? recentRank.get(right.id)! : Number.MAX_SAFE_INTEGER;
    if (leftRecent !== rightRecent) {
      return leftRecent - rightRecent;
    }
    return 0;
  });
}

export function isOtpNotifyEnabled(): boolean {
  try {
    return window.localStorage.getItem(NOTIFY_KEY) !== '0';
  } catch {
    return true;
  }
}

export function setOtpNotifyEnabled(enabled: boolean) {
  try {
    window.localStorage.setItem(NOTIFY_KEY, enabled ? '1' : '0');
  } catch {
    // ignore
  }
}

export async function ensureNotificationPermission(): Promise<NotificationPermission | 'unsupported'> {
  if (typeof window === 'undefined' || !('Notification' in window)) {
    return 'unsupported';
  }
  if (Notification.permission === 'granted' || Notification.permission === 'denied') {
    return Notification.permission;
  }
  try {
    return await Notification.requestPermission();
  } catch {
    return Notification.permission;
  }
}

export function notifyVerificationCode(payload: { title: string; body: string; code: string }) {
  if (typeof window === 'undefined' || !('Notification' in window)) {
    return;
  }
  if (!isOtpNotifyEnabled() || Notification.permission !== 'granted') {
    return;
  }
  try {
    new Notification(payload.title, {
      body: payload.body,
      tag: `otp-${payload.code}`,
    });
  } catch {
    // ignore
  }
}
