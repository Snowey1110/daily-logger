import React, { useState, useEffect, useCallback, useRef } from 'react';
import Calendar from 'react-calendar';
import 'react-calendar/dist/Calendar.css';
import { 
  Settings, 
  PenTool, 
  Calendar as CalendarIcon, 
  ArrowUpDown,
  FilePlus,
  Pencil,
  Save,
  BookOpen,
  Palette,
} from 'lucide-react';
import { cn } from '../lib/utils';
import { useReaderT } from '../readerI18n';
import { useTheme } from './ThemeProvider';
import { ThemePicker } from './ThemePicker';

interface NavigationProps {
  currentPage: number;
  totalPages: number;
  onPageJump: (page: number) => void;
  onPrev: () => void;
  onNext: () => void;
  onAction: (action: 'sketch' | 'edit') => void;
  onToggleSort: () => void;
  sortOrder: 'asc' | 'desc';
  onDateSelect: (date: Date) => void;
  isEditTextOpen: boolean;
  onSaveText: () => void;
}

const Navigation: React.FC<NavigationProps> = ({
  currentPage,
  totalPages,
  onPageJump,
  onPrev,
  onNext,
  onAction,
  onToggleSort,
  sortOrder,
  onDateSelect,
  isEditTextOpen,
  onSaveText,
}) => {
  const { t } = useReaderT();
  const { coverTheme, bgTheme, setCoverTheme, setBgTheme } = useTheme();
  const [showCalendar, setShowCalendar] = useState(false);
  const [showActionMenu, setShowActionMenu] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [jumpInput, setJumpInput] = useState('');
  const [pickerMode, setPickerMode] = useState<'cover' | 'background' | null>(null);

  const navRef = useRef<HTMLDivElement>(null);
  const accent = bgTheme.cover.accentText;
  const dropdownBg = bgTheme.colors.bg;

  const closeAll = useCallback(() => {
    setShowCalendar(false);
    setShowActionMenu(false);
    setShowSettings(false);
  }, []);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (navRef.current && !navRef.current.contains(e.target as Node)) {
        closeAll();
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [closeAll]);

  const handleJump = (e: React.FormEvent) => {
    e.preventDefault();
    const page = parseInt(jumpInput);
    if (!isNaN(page) && page >= 1 && page <= totalPages) {
      onPageJump(page);
      setJumpInput('');
    }
  };

  return (
    <>
      <div ref={navRef} className="absolute top-6 right-8 flex items-center space-x-6 z-50">
        <div className="flex items-center space-x-2 bg-white/10 backdrop-blur-md px-4 py-2 rounded-full border border-white/20">
          <span className="text-[10px] uppercase tracking-tighter opacity-70 font-sans font-semibold" style={{ color: accent }}>{t('navPage')}</span>
          <form onSubmit={handleJump} className="flex items-center">
            <input 
              type="text" 
              value={jumpInput}
              onChange={(e) => setJumpInput(e.target.value)}
              placeholder={currentPage === 0 ? '1' : String(currentPage)}
              className="bg-transparent w-8 text-center text-white border-b border-white/30 focus:outline-none font-sans text-sm pb-0.5"
            />
          </form>
          <span className="text-[10px] opacity-40 font-sans font-semibold" style={{ color: accent }}>/ {totalPages}</span>
        </div>

        <div className="flex items-center space-x-2">
          <div className="relative">
            <button 
              onClick={() => { const next = !showCalendar; closeAll(); setShowCalendar(next); }}
              className={cn(
                "p-2 hover:text-white transition-colors rounded-full hover:bg-white/5",
                showCalendar && "text-white bg-white/10"
              )}
              style={{ color: showCalendar ? undefined : accent }}
              title={t('navJumpDate')}
            >
              <CalendarIcon size={20} strokeWidth={1.5} />
            </button>
            
            {showCalendar && (
              <div
                className="absolute top-full mt-4 right-0 shadow-2xl rounded-2xl overflow-hidden border border-white/10 backdrop-blur-xl animate-in fade-in slide-in-from-top-2 duration-300 z-[100] p-2"
                style={{ backgroundColor: dropdownBg }}
              >
                <Calendar 
                  onChange={(val) => {
                    onDateSelect(val as Date);
                    setShowCalendar(false);
                  }}
                  className="!border-0 text-sm rounded-xl !bg-white/5"
                  style={{ color: accent }}
                />
              </div>
            )}
          </div>

          <div className="relative">
            <button 
              onClick={() => { const next = !showActionMenu; closeAll(); setShowActionMenu(next); }}
              className={cn(
                "p-2 hover:text-white transition-colors rounded-full hover:bg-white/5",
                showActionMenu && "text-white bg-white/10"
              )}
              style={{ color: showActionMenu ? undefined : accent }}
              title={t('navPenOptions')}
            >
              <PenTool size={20} strokeWidth={1.5} />
            </button>
            
            {showActionMenu && (
              <div
                className="absolute top-full mt-4 right-0 shadow-2xl rounded-xl overflow-hidden border border-white/10 backdrop-blur-xl animate-in fade-in slide-in-from-top-2 duration-300 z-[100] w-48 p-1"
                style={{ backgroundColor: dropdownBg }}
              >
                <button 
                  onClick={() => { onAction('sketch'); setShowActionMenu(false); }}
                  className="w-full flex items-center gap-3 px-4 py-3 text-sm font-medium hover:bg-white/10 hover:text-white transition-all rounded-lg"
                  style={{ color: accent }}
                >
                  <FilePlus size={18} />
                  <span>{t('navDrawSketch')}</span>
                </button>
                <button 
                  onClick={() => { onAction('edit'); setShowActionMenu(false); }}
                  className="w-full flex items-center gap-3 px-4 py-3 text-sm font-medium hover:bg-white/10 hover:text-white transition-all rounded-lg"
                  style={{ color: accent }}
                >
                  <Pencil size={18} />
                  <span>{t('navEditJournal')}</span>
                </button>
              </div>
            )}
          </div>

          {isEditTextOpen && (
            <button 
              onClick={onSaveText}
              className="p-2 text-green-400 hover:text-green-300 transition-colors rounded-full bg-green-500/10 border border-green-500/20"
              title={t('navSaveChanges')}
            >
              <Save size={20} strokeWidth={1.5} />
            </button>
          )}

          <button 
            onClick={onToggleSort}
            className="p-2 hover:text-white transition-colors rounded-full hover:bg-white/5"
            style={{ color: accent }}
            title={`${t('navSortPrefix')} ${sortOrder === 'asc' ? t('navSortOldest') : t('navSortNewest')}`}
          >
            <ArrowUpDown size={20} strokeWidth={1.5} className={cn(sortOrder === 'desc' && "rotate-180")} />
          </button>

          <div className="relative">
            <button
              onClick={() => { const next = !showSettings; closeAll(); setShowSettings(next); }}
              className={cn(
                "p-2 hover:text-white transition-colors rounded-full hover:bg-white/5",
                showSettings && "text-white bg-white/10"
              )}
              style={{ color: showSettings ? undefined : accent }}
              title={t('settingsTitle')}
            >
              <Settings size={20} strokeWidth={1.5} />
            </button>

            {showSettings && (
              <div
                className="absolute top-full mt-4 right-0 shadow-2xl rounded-xl overflow-hidden border border-white/10 backdrop-blur-xl animate-in fade-in slide-in-from-top-2 duration-300 z-[100] w-52 p-2"
                style={{ backgroundColor: dropdownBg }}
              >
                <h4 className="text-[10px] font-bold uppercase tracking-widest opacity-50 mb-1.5 px-2" style={{ color: accent }}>
                  {t('settingsTitle')}
                </h4>
                <button
                  onClick={() => { setPickerMode('cover'); setShowSettings(false); }}
                  className="w-full flex items-center gap-3 px-3 py-2.5 text-sm font-medium hover:bg-white/10 hover:text-white transition-all rounded-lg"
                  style={{ color: accent }}
                >
                  <BookOpen size={18} />
                  <span>{t('settingsCoverTheme')}</span>
                </button>
                <button
                  onClick={() => { setPickerMode('background'); setShowSettings(false); }}
                  className="w-full flex items-center gap-3 px-3 py-2.5 text-sm font-medium hover:bg-white/10 hover:text-white transition-all rounded-lg"
                  style={{ color: accent }}
                >
                  <Palette size={18} />
                  <span>{t('settingsBgTheme')}</span>
                </button>
              </div>
            )}
          </div>
        </div>
      </div>

      {pickerMode && (
        <ThemePicker
          mode={pickerMode}
          currentThemeId={pickerMode === 'cover' ? coverTheme.id : bgTheme.id}
          onSelect={(id) => {
            if (pickerMode === 'cover') setCoverTheme(id);
            else setBgTheme(id);
            setPickerMode(null);
          }}
          onClose={() => setPickerMode(null)}
        />
      )}
    </>
  );
};

export default Navigation;
