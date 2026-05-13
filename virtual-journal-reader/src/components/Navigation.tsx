import React, { useState } from 'react';
import Calendar from 'react-calendar';
import 'react-calendar/dist/Calendar.css';
import { 
  Settings, 
  PenTool, 
  Calendar as CalendarIcon, 
  ArrowUpDown,
  Image as ImageIcon,
  Pencil,
  Save,
  Check
} from 'lucide-react';
import { cn, type RightPageSetting } from '../lib/utils';
import { useReaderT } from '../readerI18n';

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
  rightPageSetting: RightPageSetting;
  onRightPageSettingChange: (setting: RightPageSetting) => void;
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
  rightPageSetting,
  onRightPageSettingChange,
}) => {
  const { t } = useReaderT();
  const [showCalendar, setShowCalendar] = useState(false);
  const [showActionMenu, setShowActionMenu] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [jumpInput, setJumpInput] = useState('');

  const handleJump = (e: React.FormEvent) => {
    e.preventDefault();
    const page = parseInt(jumpInput);
    if (!isNaN(page) && page >= 1 && page <= totalPages) {
      onPageJump(page);
      setJumpInput('');
    }
  };

  return (
    <div className="absolute top-6 right-8 flex items-center space-x-6 z-50">
      <div className="flex items-center space-x-2 bg-white/10 backdrop-blur-md px-4 py-2 rounded-full border border-white/20">
        <span className="text-[#d9c5b2] text-[10px] uppercase tracking-tighter opacity-70 font-sans font-semibold">{t('navPage')}</span>
        <form onSubmit={handleJump} className="flex items-center">
          <input 
            type="text" 
            value={jumpInput}
            onChange={(e) => setJumpInput(e.target.value)}
            placeholder={currentPage === 0 ? '1' : String(currentPage)}
            className="bg-transparent w-8 text-center text-white border-b border-white/30 focus:outline-none font-sans text-sm pb-0.5"
          />
        </form>
        <span className="text-[#d9c5b2] text-[10px] opacity-40 font-sans font-semibold">/ {totalPages}</span>
      </div>

      <div className="flex items-center space-x-2">
        <div className="relative">
          <button 
            onClick={() => setShowCalendar(!showCalendar)}
            className={cn(
              "p-2 text-[#d9c5b2] hover:text-white transition-colors rounded-full hover:bg-white/5",
              showCalendar && "text-white bg-white/10"
            )}
            title={t('navJumpDate')}
          >
            <CalendarIcon size={20} strokeWidth={1.5} />
          </button>
          
          {showCalendar && (
            <div className="absolute top-full mt-4 right-0 shadow-2xl rounded-2xl overflow-hidden border border-white/10 bg-[#2c1e14] backdrop-blur-xl animate-in fade-in slide-in-from-top-2 duration-300 z-[100] p-2">
              <Calendar 
                onChange={(val) => {
                  onDateSelect(val as Date);
                  setShowCalendar(false);
                }}
                className="!border-0 text-sm rounded-xl !bg-white/5 !text-[#d9c5b2]"
              />
            </div>
          )}
        </div>

        <div className="relative">
          <button 
            onClick={() => setShowActionMenu(!showActionMenu)}
            className={cn(
              "p-2 text-[#d9c5b2] hover:text-white transition-colors rounded-full hover:bg-white/5",
              showActionMenu && "text-white bg-white/10"
            )}
            title={t('navPenOptions')}
          >
            <PenTool size={20} strokeWidth={1.5} />
          </button>
          
          {showActionMenu && (
            <div className="absolute top-full mt-4 right-0 shadow-2xl rounded-xl overflow-hidden border border-white/10 bg-[#2c1e14] backdrop-blur-xl animate-in fade-in slide-in-from-top-2 duration-300 z-[100] w-48 p-1">
              <button 
                onClick={() => { onAction('sketch'); setShowActionMenu(false); }}
                className="w-full flex items-center gap-3 px-4 py-3 text-sm font-medium text-[#d9c5b2] hover:bg-white/10 hover:text-white transition-all rounded-lg"
              >
                <ImageIcon size={18} />
                <span>{t('navDrawSketch')}</span>
              </button>
              <button 
                onClick={() => { onAction('edit'); setShowActionMenu(false); }}
                className="w-full flex items-center gap-3 px-4 py-3 text-sm font-medium text-[#d9c5b2] hover:bg-white/10 hover:text-white transition-all rounded-lg"
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
          className="p-2 text-[#d9c5b2] hover:text-white transition-colors rounded-full hover:bg-white/5"
          title={`${t('navSortPrefix')} ${sortOrder === 'asc' ? t('navSortOldest') : t('navSortNewest')}`}
        >
          <ArrowUpDown size={20} strokeWidth={1.5} className={cn(sortOrder === 'desc' && "rotate-180")} />
        </button>

        <div className="relative">
          <button
            onClick={() => setShowSettings(!showSettings)}
            className={cn(
              "p-2 text-[#d9c5b2] hover:text-white transition-colors rounded-full hover:bg-white/5",
              showSettings && "text-white bg-white/10"
            )}
            title={t('settingsTitle')}
          >
            <Settings size={20} strokeWidth={1.5} />
          </button>

          {showSettings && (
            <div className="absolute top-full mt-4 right-0 shadow-2xl rounded-xl overflow-hidden border border-white/10 bg-[#2c1e14] backdrop-blur-xl animate-in fade-in slide-in-from-top-2 duration-300 z-[100] w-56 p-3">
              <h4 className="text-[10px] font-bold uppercase tracking-widest text-[#d9c5b2]/50 mb-2 px-1">
                {t('settingsRightPage')}
              </h4>
              {(['none', 'ai', 'stt'] as const).map((opt) => {
                const label = opt === 'none' ? t('settingsNone') : opt === 'ai' ? t('settingsAiRight') : t('settingsSttRight');
                return (
                  <button
                    key={opt}
                    onClick={() => { onRightPageSettingChange(opt); setShowSettings(false); }}
                    className={cn(
                      "w-full flex items-center gap-3 px-3 py-2.5 text-sm font-medium rounded-lg transition-all",
                      rightPageSetting === opt
                        ? "text-white bg-white/10"
                        : "text-[#d9c5b2] hover:bg-white/10 hover:text-white"
                    )}
                  >
                    <span className="w-4 flex-shrink-0">
                      {rightPageSetting === opt && <Check size={14} />}
                    </span>
                    <span>{label}</span>
                  </button>
                );
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default Navigation;
