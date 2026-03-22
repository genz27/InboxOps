import { create } from 'zustand';
import { persist } from 'zustand/middleware';
import { en } from './en';
import { zh } from './zh';

type Language = 'en' | 'zh';

interface I18nState {
  language: Language;
  setLanguage: (lang: Language) => void;
  t: (key: keyof typeof en) => string;
}

export const useI18n = create<I18nState>()(
  persist(
    (set, get) => ({
      language: 'zh',
      setLanguage: (lang) => set({ language: lang }),
      t: (key) => {
        const lang = get().language;
        const dict = lang === 'en' ? en : zh;
        return dict[key] || key;
      },
    }),
    {
      name: 'i18n-storage',
    }
  )
);
