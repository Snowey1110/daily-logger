import React, { createContext, useContext, useState, type ReactNode } from 'react';
import type { JournalTheme } from '../types/theme';
import { THEMES } from '../constants/themes';

interface ThemeContextType {
  coverTheme: JournalTheme;
  bgTheme: JournalTheme;
  setCoverTheme: (id: string) => void;
  setBgTheme: (id: string) => void;
}

const COVER_KEY = 'virtualJournalReader.coverTheme';
const BG_KEY = 'virtualJournalReader.bgTheme';

function loadTheme(key: string): JournalTheme {
  try {
    const id = localStorage.getItem(key);
    if (id) {
      const t = THEMES.find((th) => th.id === id);
      if (t) return t;
    }
  } catch { /* ignore */ }
  return THEMES[0];
}

const ThemeContext = createContext<ThemeContextType | undefined>(undefined);

export const ThemeProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  const [coverTheme, setCoverState] = useState<JournalTheme>(() => loadTheme(COVER_KEY));
  const [bgTheme, setBgState] = useState<JournalTheme>(() => loadTheme(BG_KEY));

  const setCoverTheme = (id: string) => {
    const t = THEMES.find((th) => th.id === id);
    if (t) {
      setCoverState(t);
      try { localStorage.setItem(COVER_KEY, id); } catch { /* ignore */ }
    }
  };

  const setBgTheme = (id: string) => {
    const t = THEMES.find((th) => th.id === id);
    if (t) {
      setBgState(t);
      try { localStorage.setItem(BG_KEY, id); } catch { /* ignore */ }
    }
  };

  return (
    <ThemeContext.Provider value={{ coverTheme, bgTheme, setCoverTheme, setBgTheme }}>
      {children}
    </ThemeContext.Provider>
  );
};

export const useTheme = () => {
  const ctx = useContext(ThemeContext);
  if (!ctx) throw new Error('useTheme must be used within ThemeProvider');
  return ctx;
};
