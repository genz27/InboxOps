import { create } from 'zustand';
import { persist } from 'zustand/middleware';

type Theme = 'light' | 'dark' | 'system';
type ResolvedTheme = 'light' | 'dark';

interface ThemeState {
  theme: Theme;
  setTheme: (theme: Theme) => void;
}

const THEME_STORAGE_KEY = 'theme-storage';
const SYSTEM_THEME_QUERY = '(prefers-color-scheme: dark)';

let systemThemeMediaQuery: MediaQueryList | null = null;
let systemThemeListenerBound = false;

function canUseThemeDom() {
  return typeof window !== 'undefined' && typeof document !== 'undefined';
}

function resolveTheme(theme: Theme): ResolvedTheme {
  if (theme === 'system') {
    if (!canUseThemeDom()) {
      return 'light';
    }
    return window.matchMedia(SYSTEM_THEME_QUERY).matches ? 'dark' : 'light';
  }
  return theme;
}

function applyTheme(theme: Theme) {
  if (!canUseThemeDom()) {
    return;
  }

  const root = document.documentElement;
  const resolvedTheme = resolveTheme(theme);

  root.classList.toggle('dark', resolvedTheme === 'dark');
  root.style.colorScheme = resolvedTheme;
  root.dataset.theme = theme;
}

export const useTheme = create<ThemeState>()(
  persist(
    (set) => ({
      theme: 'system',
      setTheme: (theme) => {
        set({ theme });
        applyTheme(theme);
      },
    }),
    {
      name: THEME_STORAGE_KEY,
      onRehydrateStorage: () => (state) => {
        if (state) {
          applyTheme(state.theme);
        }
      },
    }
  )
);

function bindSystemThemeListener() {
  if (!canUseThemeDom() || systemThemeListenerBound) {
    return;
  }

  systemThemeMediaQuery = window.matchMedia(SYSTEM_THEME_QUERY);

  const handleSystemThemeChange = () => {
    if (useTheme.getState().theme === 'system') {
      applyTheme('system');
    }
  };

  if (typeof systemThemeMediaQuery.addEventListener === 'function') {
    systemThemeMediaQuery.addEventListener('change', handleSystemThemeChange);
  } else {
    systemThemeMediaQuery.addListener(handleSystemThemeChange);
  }

  systemThemeListenerBound = true;
}

if (canUseThemeDom()) {
  bindSystemThemeListener();
  applyTheme(useTheme.getState().theme);
}
