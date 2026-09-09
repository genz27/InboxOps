import { create } from 'zustand';
import { persist } from 'zustand/middleware';

type Theme = 'dark';

interface ThemeState {
  theme: Theme;
  setTheme: (theme: Theme) => void;
}

const THEME_STORAGE_KEY = 'theme-storage';

function canUseThemeDom() {
  return typeof window !== 'undefined' && typeof document !== 'undefined';
}

function applyDark() {
  if (!canUseThemeDom()) {
    return;
  }
  const root = document.documentElement;
  root.classList.add('dark');
  root.classList.remove('light');
  root.style.colorScheme = 'dark';
  root.dataset.theme = 'dark';
}

export const useTheme = create<ThemeState>()(
  persist(
    (set) => ({
      theme: 'dark',
      setTheme: () => {
        set({ theme: 'dark' });
        applyDark();
      },
    }),
    {
      name: THEME_STORAGE_KEY,
      version: 4,
      migrate: () => ({ theme: 'dark' as const }),
      onRehydrateStorage: () => () => {
        applyDark();
      },
    },
  ),
);

if (canUseThemeDom()) {
  applyDark();
}
