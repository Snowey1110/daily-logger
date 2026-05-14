import React, { createContext, useContext, useState, useEffect, type ReactNode } from 'react';
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

function findTheme(id: string | null | undefined): JournalTheme | undefined {
  if (!id) return undefined;
  return THEMES.find((th) => th.id === id);
}

function loadTheme(key: string): JournalTheme {
  try {
    const id = localStorage.getItem(key);
    const t = findTheme(id);
    if (t) return t;
  } catch { /* ignore */ }
  return THEMES[0];
}

function persistSettings(patch: Record<string, string>) {
  fetch('/api/reader-settings', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(patch),
  }).catch(() => {});
}

const ThemeContext = createContext<ThemeContextType | undefined>(undefined);

export const ThemeProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  const [coverTheme, setCoverState] = useState<JournalTheme>(() => loadTheme(COVER_KEY));
  const [bgTheme, setBgState] = useState<JournalTheme>(() => loadTheme(BG_KEY));

  useEffect(() => {
    fetch('/api/reader-settings')
      .then((r) => r.json())
      .then((data) => {
        const ct = findTheme(data.coverTheme);
        const bt = findTheme(data.bgTheme);
        if (ct) { setCoverState(ct); try { localStorage.setItem(COVER_KEY, ct.id); } catch {} }
        if (bt) { setBgState(bt); try { localStorage.setItem(BG_KEY, bt.id); } catch {} }
      })
      .catch(() => {});
  }, []);

  const setCoverTheme = (id: string) => {
    const t = THEMES.find((th) => th.id === id);
    if (t) {
      setCoverState(t);
      try { localStorage.setItem(COVER_KEY, id); } catch { /* ignore */ }
      persistSettings({ coverTheme: id });
    }
  };

  const setBgTheme = (id: string) => {
    const t = THEMES.find((th) => th.id === id);
    if (t) {
      setBgState(t);
      try { localStorage.setItem(BG_KEY, id); } catch { /* ignore */ }
      persistSettings({ bgTheme: id });
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
