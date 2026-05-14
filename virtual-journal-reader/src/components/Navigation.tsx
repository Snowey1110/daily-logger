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
  Wifi,
  WifiOff,
  Copy,
  Check,
  Smartphone,
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
  isMobile?: boolean;
  singlePageMode?: boolean;
  onToggleSinglePage?: () => void;
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
  isMobile,
  singlePageMode,
  onToggleSinglePage,
}) => {
  const { t } = useReaderT();
  const { coverTheme, bgTheme, setCoverTheme, setBgTheme } = useTheme();
  const [showCalendar, setShowCalendar] = useState(false);
  const [showActionMenu, setShowActionMenu] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [jumpInput, setJumpInput] = useState('');
  const [pickerMode, setPickerMode] = useState<'cover' | 'background' | null>(null);

  const [lanEnabled, setLanEnabled] = useState(false);
  const [lanIp, setLanIp] = useState('');
  const [lanPort, setLanPort] = useState(8765);
  const [lanLoading, setLanLoading] = useState(false);
  const [copied, setCopied] = useState(false);

  const navRef = useRef<HTMLDivElement>(null);
  const accent = bgTheme.cover.accentText;
  const dropdownBg = bgTheme.colors.bg;

  useEffect(() => {
    fetch('/api/lan-status')
      .then((r) => r.json())
      .then((d) => { setLanEnabled(d.enabled); setLanIp(d.ip || ''); setLanPort(d.port || 8765); })
      .catch(() => {});
  }, []);

  const toggleLan = async () => {
    setLanLoading(true);
    try {
      const res = await fetch('/api/lan-toggle', { method: 'POST' });
      const data = await res.json();
      if (data.ok) {
        setLanEnabled(data.enabled);
        setLanIp(data.ip || '');
        setLanPort(data.port || 8765);
      }
    } catch { /* ignore */ }
    setLanLoading(false);
  };

  const copyLanUrl = () => {
    const url = `http://${lanIp}:${lanPort}/`;
    navigator.clipboard.writeText(url).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    });
  };

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
      <div ref={navRef} className={cn(
        "absolute z-50 flex items-center",
        isMobile
          ? "top-2 right-2 left-2 justify-between space-x-2"
          : "top-6 right-8 space-x-6"
      )}>
        <div className={cn(
          "flex items-center space-x-2 bg-white/10 backdrop-blur-md rounded-full border border-white/20",
          isMobile ? "px-3 py-1.5" : "px-4 py-2"
        )}>
          <span className="text-[10px] uppercase tracking-tighter opacity-70 font-sans font-semibold" style={{ color: accent }}>{t('navPage')}</span>
          <form onSubmit={handleJump} className="flex items-center">
            <input 
              type="text" 
              value={jumpInput}
              onChange={(e) => setJumpInput(e.target.value)}
              placeholder={currentPage === 0 ? '1' : String(currentPage)}
              className={cn(
                "bg-transparent text-center text-white border-b border-white/30 focus:outline-none font-sans pb-0.5",
                isMobile ? "w-6 text-xs min-h-[44px]" : "w-8 text-sm"
              )}
            />
          </form>
          <span className="text-[10px] opacity-40 font-sans font-semibold" style={{ color: accent }}>/ {totalPages}</span>
        </div>

        <div className="flex items-center space-x-1 md:space-x-2">
          <div className="relative">
            <button 
              onClick={() => { const next = !showCalendar; closeAll(); setShowCalendar(next); }}
              className={cn(
                "p-2 hover:text-white transition-colors rounded-full hover:bg-white/5 min-h-[44px] min-w-[44px] flex items-center justify-center",
                showCalendar && "text-white bg-white/10"
              )}
              style={{ color: showCalendar ? undefined : accent }}
              title={t('navJumpDate')}
            >
              <CalendarIcon size={isMobile ? 18 : 20} strokeWidth={1.5} />
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
                "p-2 hover:text-white transition-colors rounded-full hover:bg-white/5 min-h-[44px] min-w-[44px] flex items-center justify-center",
                showActionMenu && "text-white bg-white/10"
              )}
              style={{ color: showActionMenu ? undefined : accent }}
              title={t('navPenOptions')}
            >
              <PenTool size={isMobile ? 18 : 20} strokeWidth={1.5} />
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
            className="p-2 hover:text-white transition-colors rounded-full hover:bg-white/5 min-h-[44px] min-w-[44px] flex items-center justify-center"
            style={{ color: accent }}
            title={`${t('navSortPrefix')} ${sortOrder === 'asc' ? t('navSortOldest') : t('navSortNewest')}`}
          >
            <ArrowUpDown size={isMobile ? 18 : 20} strokeWidth={1.5} className={cn(sortOrder === 'desc' && "rotate-180")} />
          </button>

          <div className="relative">
            <button
              onClick={() => { const next = !showSettings; closeAll(); setShowSettings(next); }}
              className={cn(
                "p-2 hover:text-white transition-colors rounded-full hover:bg-white/5 min-h-[44px] min-w-[44px] flex items-center justify-center",
                showSettings && "text-white bg-white/10"
              )}
              style={{ color: showSettings ? undefined : accent }}
              title={t('settingsTitle')}
            >
              <Settings size={isMobile ? 18 : 20} strokeWidth={1.5} />
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

                {onToggleSinglePage && (
                  <button
                    onClick={() => { onToggleSinglePage(); setShowSettings(false); }}
                    className="w-full flex items-center gap-3 px-3 py-2.5 text-sm font-medium hover:bg-white/10 hover:text-white transition-all rounded-lg"
                    style={{ color: singlePageMode ? '#22c55e' : accent }}
                  >
                    <Smartphone size={18} />
                    <span className="flex-1 text-left">Single Page</span>
                    <span
                      className="w-8 h-4 rounded-full relative transition-colors flex-shrink-0"
                      style={{ backgroundColor: singlePageMode ? '#22c55e' : `${accent}30` }}
                    >
                      <span
                        className="absolute top-0.5 w-3 h-3 rounded-full bg-white shadow transition-all"
                        style={{ left: singlePageMode ? '1rem' : '0.125rem' }}
                      />
                    </span>
                  </button>
                )}

                <div className="mt-1 pt-1" style={{ borderTop: `1px solid ${accent}20` }}>
                  <button
                    onClick={toggleLan}
                    disabled={lanLoading}
                    className="w-full flex items-center gap-3 px-3 py-2.5 text-sm font-medium hover:bg-white/10 hover:text-white transition-all rounded-lg"
                    style={{ color: lanEnabled ? '#22c55e' : accent }}
                  >
                    {lanEnabled ? <Wifi size={18} /> : <WifiOff size={18} />}
                    <span className="flex-1 text-left">
                      {lanEnabled ? 'Phone Access On' : 'Phone Access Off'}
                    </span>
                    <span
                      className="w-8 h-4 rounded-full relative transition-colors flex-shrink-0"
                      style={{ backgroundColor: lanEnabled ? '#22c55e' : `${accent}30` }}
                    >
                      <span
                        className="absolute top-0.5 w-3 h-3 rounded-full bg-white shadow transition-all"
                        style={{ left: lanEnabled ? '1rem' : '0.125rem' }}
                      />
                    </span>
                  </button>
                  {lanEnabled && lanIp && (
                    <div className="px-3 pb-2">
                      <div className="flex items-center gap-1.5 mt-1">
                        <code className="text-[10px] px-2 py-1 rounded bg-white/5 flex-1 truncate" style={{ color: accent }}>
                          {lanIp}:{lanPort}
                        </code>
                        <button
                          onClick={copyLanUrl}
                          className="p-1 rounded hover:bg-white/10 transition-colors flex-shrink-0"
                          style={{ color: accent }}
                          title="Copy URL"
                        >
                          {copied ? <Check size={14} className="text-green-400" /> : <Copy size={14} />}
                        </button>
                      </div>
                      <span className="text-[9px] opacity-40 mt-0.5 block" style={{ color: accent }}>
                        Open this on your phone
                      </span>
                    </div>
                  )}
                </div>
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
