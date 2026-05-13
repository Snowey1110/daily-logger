import React from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { X, Check, Book as BookIcon } from 'lucide-react';
import { THEMES } from '../constants/themes';
import type { JournalTheme } from '../types/theme';
import { useReaderT } from '../readerI18n';
import { useTheme } from './ThemeProvider';

interface ThemePickerProps {
  mode: 'cover' | 'background';
  currentThemeId: string;
  onSelect: (id: string) => void;
  onClose: () => void;
}

const CoverPreview: React.FC<{ theme: JournalTheme; selected: boolean }> = ({ theme, selected }) => (
  <div
    className="relative w-full aspect-[1.4/1] rounded-lg overflow-hidden border-2 transition-all"
    style={{
      backgroundImage: theme.cover.gradient,
      borderColor: selected ? theme.cover.accentText : 'transparent',
      boxShadow: selected ? `0 0 0 2px ${theme.cover.accentText}40` : 'none',
    }}
  >
    <div className="absolute inset-0 opacity-15 pointer-events-none bg-[url('https://www.transparenttextures.com/patterns/leather.png')]" />
    <div
      className="absolute inset-2 border rounded pointer-events-none"
      style={{ borderColor: theme.cover.accentBorder }}
    />
    <div className="flex flex-col items-center justify-center h-full gap-2 px-3">
      <div
        className="w-8 h-8 rounded-full flex items-center justify-center border"
        style={{ backgroundColor: theme.cover.accentBg, borderColor: theme.cover.accentBorder }}
      >
        <BookIcon size={14} style={{ color: theme.cover.accentText }} />
      </div>
      <span
        className="text-[9px] font-serif font-bold tracking-widest uppercase leading-tight text-center"
        style={{ color: theme.cover.accentText }}
      >
        Daily Logger
      </span>
      <div className="h-px w-8" style={{ backgroundColor: theme.cover.accentBorder }} />
      <span
        className="text-[7px] font-serif italic tracking-wider"
        style={{ color: theme.cover.subtitleText }}
      >
        Journal
      </span>
    </div>
    {selected && (
      <div className="absolute top-1.5 right-1.5 w-5 h-5 rounded-full flex items-center justify-center" style={{ backgroundColor: theme.cover.accentText }}>
        <Check size={12} className="text-white" />
      </div>
    )}
  </div>
);

const BackgroundPreview: React.FC<{ theme: JournalTheme; selected: boolean }> = ({ theme, selected }) => (
  <div
    className="relative w-full aspect-[1.4/1] rounded-lg overflow-hidden border-2 transition-all flex"
    style={{
      backgroundColor: theme.colors.bg,
      borderColor: selected ? (theme.cover.isDark ? '#818cf8' : '#92400e') : 'transparent',
      boxShadow: selected ? `0 0 0 2px ${theme.cover.isDark ? 'rgba(129,140,248,0.25)' : 'rgba(146,64,14,0.25)'}` : 'none',
    }}
  >
    <div className="flex-1 flex items-center justify-center p-1.5">
      <div
        className="w-full h-full rounded-sm flex overflow-hidden"
        style={{ backgroundColor: theme.colors.bookInner }}
      >
        {/* Left page */}
        <div className="flex-1 p-2 flex flex-col gap-1" style={{ borderRight: `1px solid ${theme.colors.border}` }}>
          <div className="h-1 w-8 rounded-full opacity-40" style={{ backgroundColor: theme.colors.text }} />
          <div className="h-0.5 w-full rounded-full opacity-15" style={{ backgroundColor: theme.colors.text }} />
          <div className="h-0.5 w-10 rounded-full opacity-15" style={{ backgroundColor: theme.colors.text }} />
          <div className="h-0.5 w-full rounded-full opacity-10" style={{ backgroundColor: theme.colors.text }} />
        </div>
        {/* Right page */}
        <div className="flex-1 p-2 flex flex-col gap-1">
          <div className="h-1 w-6 rounded-full opacity-30" style={{ backgroundColor: theme.colors.text }} />
          <div className="h-0.5 w-full rounded-full opacity-10" style={{ backgroundColor: theme.colors.text }} />
          <div className="h-0.5 w-8 rounded-full opacity-10" style={{ backgroundColor: theme.colors.text }} />
        </div>
        {/* Bookmark tabs */}
        <div className="flex flex-col gap-0.5 py-2 pr-0.5">
          <div className="w-1.5 h-3 rounded-r-sm" style={{ backgroundColor: theme.colors.tabs.journal.active }} />
          <div className="w-1 h-2.5 rounded-r-sm" style={{ backgroundColor: theme.colors.tabs.stt.bg }} />
          <div className="w-1 h-2.5 rounded-r-sm" style={{ backgroundColor: theme.colors.tabs.ai.bg }} />
        </div>
      </div>
    </div>
    {selected && (
      <div
        className="absolute top-1.5 right-1.5 w-5 h-5 rounded-full flex items-center justify-center"
        style={{ backgroundColor: theme.cover.isDark ? '#818cf8' : '#92400e' }}
      >
        <Check size={12} className="text-white" />
      </div>
    )}
  </div>
);

export const ThemePicker: React.FC<ThemePickerProps> = ({ mode, currentThemeId, onSelect, onClose }) => {
  const { t } = useReaderT();
  const { bgTheme } = useTheme();
  const title = mode === 'cover' ? t('themePickerCover') : t('themePickerBackground');
  const accent = bgTheme.cover.accentText;

  return (
    <AnimatePresence>
      <motion.div
        className="fixed inset-0 z-[200] flex items-center justify-center bg-black/60 backdrop-blur-sm"
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onClose}
      >
        <motion.div
          className="relative w-full max-w-lg rounded-2xl border border-white/10 shadow-2xl p-6"
          style={{ backgroundColor: bgTheme.colors.bg }}
          initial={{ scale: 0.9, opacity: 0 }}
          animate={{ scale: 1, opacity: 1 }}
          exit={{ scale: 0.9, opacity: 0 }}
          transition={{ type: 'spring', stiffness: 300, damping: 25 }}
          onClick={(e) => e.stopPropagation()}
        >
          <div className="flex items-center justify-between mb-5">
            <h2 className="text-sm font-bold uppercase tracking-widest font-sans opacity-80" style={{ color: accent }}>
              {title}
            </h2>
            <button
              onClick={onClose}
              className="p-1.5 rounded-full hover:bg-white/5 transition-colors opacity-50 hover:opacity-100"
              style={{ color: accent }}
            >
              <X size={18} />
            </button>
          </div>

          <div className="grid grid-cols-2 gap-4">
            {THEMES.map((theme) => {
              const isSelected = theme.id === currentThemeId;
              return (
                <button
                  key={theme.id}
                  onClick={() => onSelect(theme.id)}
                  className="group flex flex-col gap-2 text-left focus:outline-none"
                >
                  {mode === 'cover' ? (
                    <CoverPreview theme={theme} selected={isSelected} />
                  ) : (
                    <BackgroundPreview theme={theme} selected={isSelected} />
                  )}
                  <span
                    className="text-xs font-sans font-medium tracking-wide text-center w-full transition-colors"
                    style={{ color: accent, opacity: isSelected ? 1 : 0.5 }}
                  >
                    {theme.name}
                  </span>
                </button>
              );
            })}
          </div>
        </motion.div>
      </motion.div>
    </AnimatePresence>
  );
};
