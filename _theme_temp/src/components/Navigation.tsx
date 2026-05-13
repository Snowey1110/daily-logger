import React, { useState } from 'react';
import Calendar from 'react-calendar';
import 'react-calendar/dist/Calendar.css';
import { 
  Settings, 
  PenTool, 
  Calendar as CalendarIcon, 
  ChevronLeft, 
  ChevronRight,
  ArrowUpDown,
  Search,
  X,
  Check
} from 'lucide-react';
import { cn } from '../lib/utils';
import { useTheme } from './ThemeProvider';
import { THEMES } from '../constants/themes';

interface NavigationProps {
  currentPage: number;
  totalPages: number;
  onPageJump: (page: number) => void;
  onPrev: () => void;
  onNext: () => void;
  onOpenSketch: () => void;
  onToggleSort: () => void;
  sortOrder: 'asc' | 'desc';
  onDateSelect: (date: Date) => void;
}

const Navigation: React.FC<NavigationProps> = ({
  currentPage,
  totalPages,
  onPageJump,
  onPrev,
  onNext,
  onOpenSketch,
  onToggleSort,
  sortOrder,
  onDateSelect
}) => {
  const { currentTheme, setTheme } = useTheme();
  const [showCalendar, setShowCalendar] = useState(false);
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
    <div className={cn("flex items-center justify-between px-6 py-3 border-b shadow-sm z-[50] sticky top-0 transition-colors duration-500", currentTheme.colors.nav, "backdrop-blur-md", currentTheme.colors.border)}>
      <div className="flex items-center gap-3">
        <button 
          onClick={onPrev}
          disabled={currentPage <= 1}
          className={cn("p-2 rounded-lg disabled:opacity-30 transition-colors", currentTheme.id === 'parchment' ? "hover:bg-zinc-100" : "hover:bg-white/10")}
        >
          <ChevronLeft size={20} className={currentTheme.colors.text} />
        </button>
        <form onSubmit={handleJump} className="flex items-center gap-2">
          <span className={cn("text-sm font-medium uppercase tracking-tighter opacity-50", currentTheme.colors.text)}>Page</span>
          <input 
            type="text" 
            value={jumpInput}
            onChange={(e) => setJumpInput(e.target.value)}
            placeholder={currentPage.toString()}
            className={cn("w-12 h-8 text-center border rounded-md font-serif focus:outline-none focus:ring-2 focus:ring-indigo-500/20 bg-transparent", currentTheme.colors.text, currentTheme.colors.border)}
          />
          <span className={cn("text-sm font-medium uppercase tracking-tighter opacity-50", currentTheme.colors.text)}>of {totalPages}</span>
        </form>
        <button 
          onClick={onNext}
          disabled={currentPage >= totalPages}
          className={cn("p-2 rounded-lg disabled:opacity-30 transition-colors", currentTheme.id === 'parchment' ? "hover:bg-zinc-100" : "hover:bg-white/10")}
        >
          <ChevronRight size={20} className={currentTheme.colors.text} />
        </button>
      </div>

      <div className="flex items-center gap-2">
        <div className="relative">
          <button 
            onClick={() => { setShowCalendar(!showCalendar); setShowSettings(false); }}
            className={cn(
              "p-2 rounded-lg transition-all flex items-center gap-2 px-3",
              showCalendar 
                ? "bg-indigo-500/10 text-indigo-500 ring-1 ring-indigo-500/30" 
                : cn("hover:bg-zinc-100 text-zinc-600", currentTheme.colors.text)
            )}
          >
            <CalendarIcon size={18} />
            <span className="text-sm font-medium">Jump to Date</span>
          </button>
          
          {showCalendar && (
            <div className={cn("absolute top-full mt-2 right-0 shadow-2xl rounded-xl overflow-hidden border animate-in fade-in slide-in-from-top-2 duration-200 z-[60]", currentTheme.colors.bookInner, currentTheme.colors.border)}>
               <div className={cn("p-2 border-b flex justify-end", currentTheme.colors.border)}>
                <button onClick={() => setShowCalendar(false)} className="p-1 hover:bg-zinc-100 rounded">
                  <X size={16} className={currentTheme.colors.text} />
                </button>
              </div>
              <Calendar 
                onChange={(val) => {
                  onDateSelect(val as Date);
                  setShowCalendar(false);
                }}
                className={cn("!border-0 text-sm !bg-transparent", currentTheme.colors.text)}
              />
            </div>
          )}
        </div>

        <button 
          onClick={onOpenSketch}
          className={cn("p-2 rounded-lg transition-all flex items-center gap-2 px-3", currentTheme.colors.text, "hover:bg-indigo-500/10 hover:text-indigo-500")}
          title="Sketch/Note"
        >
          <PenTool size={18} />
          <span className="text-sm font-medium">Sketch</span>
        </button>

        <button 
          onClick={onToggleSort}
          className={cn("p-2 rounded-lg transition-all flex items-center gap-2 px-3", currentTheme.colors.text, "hover:bg-white/10")}
          title="Sort Order"
        >
          <ArrowUpDown size={18} className={cn(sortOrder === 'desc' && "rotate-180")} />
          <span className="text-sm font-medium">{sortOrder === 'asc' ? 'Oldest' : 'Newest'}</span>
        </button>

        <div className={cn("h-6 w-[1px] mx-2", currentTheme.colors.border, "opacity-50")} />

        <div className="relative">
          <button 
            onClick={() => { setShowSettings(!showSettings); setShowCalendar(false); }}
            className={cn(
              "p-2 rounded-lg transition-all",
              showSettings 
                ? "bg-indigo-500/10 text-indigo-500 ring-1 ring-indigo-500/30" 
                : cn("hover:bg-zinc-100 text-zinc-600", currentTheme.colors.text)
            )}
          >
            <Settings size={18} />
          </button>

          {showSettings && (
            <div className={cn(
              "absolute top-full mt-2 right-0 w-64 shadow-2xl rounded-xl overflow-hidden border p-4 animate-in fade-in slide-in-from-top-2 duration-200 z-[60]",
              currentTheme.colors.bookInner, 
              currentTheme.colors.border
            )}>
              <h3 className={cn("text-xs font-bold uppercase tracking-widest mb-4 opacity-50", currentTheme.colors.text)}>Select Theme</h3>
              <div className="grid grid-cols-1 gap-2">
                {THEMES.map(theme => (
                  <button
                    key={theme.id}
                    onClick={() => {
                      setTheme(theme.id);
                      setShowSettings(false);
                    }}
                    className={cn(
                      "flex items-center justify-between p-3 rounded-lg group transition-all border",
                      currentTheme.id === theme.id 
                        ? "border-indigo-500 bg-indigo-500/10 shadow-sm" 
                        : cn("border-transparent hover:bg-white/5", currentTheme.colors.border)
                    )}
                  >
                    <div className="flex items-center gap-3">
                      <div className={cn("w-4 h-4 rounded-full border shadow-sm", theme.colors.bg)} />
                      <span className={cn("text-sm font-medium", currentTheme.colors.text)}>{theme.name}</span>
                    </div>
                    {currentTheme.id === theme.id && (
                      <Check size={14} className="text-indigo-500" />
                    )}
                  </button>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default Navigation;
