export type ShortcutAction =
  | 'nextMessage'
  | 'prevMessage'
  | 'copyOtp'
  | 'archive'
  | 'focusSearch'
  | 'compose'
  | 'refresh'
  | 'toggleUnreadFilter';

/** 输入框内不拦截快捷键 */
export function isEditableTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) {
    return false;
  }
  const tag = target.tagName.toLowerCase();
  if (tag === 'input' || tag === 'textarea' || tag === 'select') {
    return true;
  }
  return target.isContentEditable;
}

export function resolveShortcut(event: KeyboardEvent): ShortcutAction | null {
  if (event.ctrlKey || event.metaKey || event.altKey) {
    return null;
  }
  const key = event.key.toLowerCase();
  if (key === 'j') return 'nextMessage';
  if (key === 'k') return 'prevMessage';
  if (key === 'c') return 'copyOtp';
  if (key === 'e') return 'archive';
  if (key === '/') return 'focusSearch';
  if (key === 'n') return 'compose';
  if (key === 'r') return 'refresh';
  if (key === 'u') return 'toggleUnreadFilter';
  return null;
}
