import React, { useState, useEffect, useCallback } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { cn, JournalEntry, JournalSection } from '../lib/utils';
import { useTheme } from './ThemeProvider';
import { JournalTheme } from '../types/theme';
import Cover from './Cover';
import Navigation from './Navigation';
import { DrawingCanvas } from './DrawingCanvas';
import { MOCK_ENTRIES } from '../data/mockEntries';
import { Clock, MessageSquare, BrainCircuit, Type } from 'lucide-react';

const JournalBook: React.FC = () => {
  const { currentTheme } = useTheme();
  const [entries, setEntries] = useState<JournalEntry[]>(MOCK_ENTRIES);
  const [isOpen, setIsOpen] = useState(false);
  const [currentPage, setCurrentPage] = useState(0); // 0 is cover
  const [activeSection, setActiveSection] = useState<JournalSection>('journal');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('asc');
  const [isSketching, setIsSketching] = useState(false);

  // Group entries into pairs of pages (left, right)
  const sortedEntries = [...entries].sort((a, b) => {
    const timeA = new Date(a.date).getTime();
    const timeB = new Date(b.date).getTime();
    return sortOrder === 'asc' ? timeA - timeB : timeB - timeA;
  });

  const totalPages = Math.ceil(sortedEntries.length / 2) + 1; // +1 for cover

  const handleNext = useCallback(() => {
    if (currentPage < totalPages - 1) setCurrentPage(prev => prev + 1);
  }, [currentPage, totalPages]);

  const handlePrev = useCallback(() => {
    if (currentPage > 0) setCurrentPage(prev => prev - 1);
  }, [currentPage]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const key = e.key.toLowerCase();
      if (key === 'd' || key === 'arrowright') handleNext();
      if (key === 'a' || key === 'arrowleft') handlePrev();
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleNext, handlePrev]);

  // Jump to page by date
  const handleDateSelect = (date: Date) => {
    const dateStr = date.toLocaleDateString('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric'
    });
    const index = sortedEntries.findIndex(e => e.date === dateStr);
    if (index !== -1) {
      setIsOpen(true);
      setCurrentPage(Math.floor(index / 2) + 1);
    }
  };

  const saveSketch = (dataUrl: string) => {
    // For demo, we just update the current entry if in book view
    if (currentPage > 0) {
      const entryIdx = (currentPage - 1) * 2;
      const updatedEntries = [...entries];
      const targetEntry = sortedEntries[entryIdx];
      const realIdx = entries.findIndex(e => e.id === targetEntry.id);
      if (realIdx !== -1) {
        updatedEntries[realIdx] = { ...updatedEntries[realIdx], sketch: dataUrl };
        setEntries(updatedEntries);
      }
    }
    setIsSketching(false);
  };

  const getPageContent = (index: number) => {
    const entryIndex = (currentPage - 1) * 2 + index;
    return sortedEntries[entryIndex];
  };

  return (
    <div className={cn("min-h-screen select-none overflow-hidden flex flex-col transition-colors duration-500", currentTheme.colors.bg, currentTheme.colors.text)}>
      <Navigation 
        currentPage={currentPage}
        totalPages={totalPages}
        onPageJump={setCurrentPage}
        onPrev={handlePrev}
        onNext={handleNext}
        onOpenSketch={() => setIsSketching(true)}
        onToggleSort={() => setSortOrder(prev => prev === 'asc' ? 'desc' : 'asc')}
        sortOrder={sortOrder}
        onDateSelect={handleDateSelect}
      />

      <div className="flex-1 flex items-center justify-center p-8 md:p-16 perspective-[2000px]">
        <div className="relative w-full max-w-6xl aspect-[1.4/1] flex shadow-2xl rounded-xl">
          
          {/* Background shadows for book depth */}
          <div className="absolute inset-0 bg-zinc-300 -z-10 rounded-xl translate-x-1 translate-y-1" />
          <div className="absolute inset-0 bg-zinc-200 -z-20 rounded-xl translate-x-2 translate-y-2 opacity-50" />

          {/* Book Spine */}
          <div className={cn(
             "absolute left-1/2 -ml-1 top-0 bottom-0 w-2 backdrop-blur z-20 shadow-[0_0_15px_rgba(0,0,0,0.2)] transition-colors duration-500",
             currentTheme.colors.spine,
             currentPage === 0 && "hidden"
          )} />

          <AnimatePresence mode="wait">
            {currentPage === 0 ? (
              <motion.div 
                key="cover"
                className="w-full h-full"
                initial={{ rotateY: -10 }}
                animate={{ rotateY: 0 }}
                exit={{ rotateY: -90, opacity: 0 }}
                transition={{ duration: 0.8, ease: "easeInOut" }}
                style={{ transformOrigin: 'left center' }}
              >
                <Cover title="Daily Logger" theme={currentTheme} onClick={() => { setIsOpen(true); setCurrentPage(1); }} />
              </motion.div>
            ) : (
              <motion.div 
                key={`page-pair-${currentPage}`}
                className={cn("flex w-full h-full overflow-hidden rounded-xl border shadow-inner relative transition-colors duration-500", currentTheme.colors.bookInner, currentTheme.colors.border)}
                initial={{ scale: 0.95, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                transition={{ duration: 0.5 }}
              >
                {/* Book Texture */}
                <div className="absolute inset-0 opacity-[0.03] pointer-events-none bg-[url('https://www.transparenttextures.com/patterns/handmade-paper.png')]" />

                {/* Left Page */}
                <Page 
                  entry={getPageContent(0)} 
                  section={activeSection}
                  side="left"
                  theme={currentTheme}
                />

                {/* Right Page */}
                <Page 
                  entry={getPageContent(1)} 
                  section={activeSection}
                  side="right"
                  theme={currentTheme}
                />

                {/* Bookmark Tabs */}
                <div className="absolute -right-2 top-20 flex flex-col gap-2 z-30">
                  <BookmarkTab 
                    label="Journal" 
                    icon={<Type size={16} />}
                    isActive={activeSection === 'journal'} 
                    onClick={() => setActiveSection('journal')}
                    color={currentTheme.colors.tabs.journal.bg}
                    activeColor={currentTheme.colors.tabs.journal.active}
                    theme={currentTheme}
                  />
                  <BookmarkTab 
                    label="STT" 
                    icon={<MessageSquare size={16} />}
                    isActive={activeSection === 'stt'} 
                    onClick={() => setActiveSection('stt')}
                    color={currentTheme.colors.tabs.stt.bg}
                    activeColor={currentTheme.colors.tabs.stt.active}
                    theme={currentTheme}
                  />
                  <BookmarkTab 
                    label="AI Report" 
                    icon={<BrainCircuit size={16} />}
                    isActive={activeSection === 'ai'} 
                    onClick={() => setActiveSection('ai')}
                    color={currentTheme.colors.tabs.ai.bg}
                    activeColor={currentTheme.colors.tabs.ai.active}
                    theme={currentTheme}
                  />
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </div>

      {isSketching && (
        <DrawingCanvas 
          onSave={saveSketch} 
          onClose={() => setIsSketching(false)} 
        />
      )}
      
      <div className="text-center pb-4 text-zinc-400 text-xs font-medium tracking-widest uppercase">
        Use A/D or Arrows to flip · Drag edges to flip
      </div>
    </div>
  );
};

const Page: React.FC<{ entry?: JournalEntry; section: JournalSection; side: 'left' | 'right'; theme: JournalTheme }> = ({ entry, section, side, theme }) => {
  if (!entry) {
    return (
      <div className={cn(
        "flex-1 h-full p-12 flex items-center justify-center transition-colors duration-500",
        side === 'left' ? "border-r" : "border-l",
        theme.colors.border
      )}>
        <p className={cn("italic font-serif tracking-widest uppercase text-xs opacity-40 transition-colors duration-500", theme.colors.textMuted)}>
          The rest is still unwritten
        </p>
      </div>
    );
  }

  const getContent = () => {
    switch(section) {
      case 'stt': return entry.speechToText;
      case 'ai': return entry.aiReport;
      default: return entry.journal;
    }
  };

  const getSectionTitle = () => {
    switch(section) {
      case 'stt': return 'Speech to Text Transcript';
      case 'ai': return 'AI Intelligence Report';
      default: return 'Journal Entry';
    }
  };

  return (
    <div className={cn(
      "flex-1 h-full flex flex-col p-12 relative overflow-hidden transition-colors duration-500",
      side === 'left' ? cn("border-r", theme.colors.border) : cn("border-l shadow-[-10px_0_10px_rgba(0,0,0,0.02)]", theme.colors.border)
    )}>
      <div className="flex items-center justify-between mb-8 opacity-60">
        <div className="flex items-center gap-2 font-serif">
          <Clock size={14} className={theme.colors.textMuted} />
          <span className={cn("text-xs uppercase tracking-widest font-bold transition-colors duration-500", theme.colors.textMuted)}>{entry.time}</span>
        </div>
        <span className={cn("text-xs font-serif italic transition-colors duration-500", theme.colors.textMuted)}>{entry.date}</span>
      </div>

      <h2 className={cn("text-lg font-serif font-bold mb-6 border-b pb-2 transition-colors duration-500", theme.colors.text, theme.colors.border, "opacity-80")}>
        {getSectionTitle()}
      </h2>

      <div className={cn("flex-1 prose prose-p:font-serif prose-p:leading-relaxed prose-p:text-lg transition-colors duration-500", theme.colors.text)}>
        {getContent().split('\n').map((para, i) => (
          <p key={i} className="mb-4 opacity-90">{para}</p>
        ))}
      </div>

      {entry.sketch && (
        <div className={cn("mt-8 border rounded-lg p-2 rotate-1 shadow-sm max-h-40 overflow-hidden transition-colors duration-500", theme.colors.border, "bg-white/10")}>
          <img src={entry.sketch} alt="Entry sketch" className="w-full grayscale brightness-90 hover:grayscale-0 transition-all opacity-80" />
        </div>
      )}

      {/* Page Number */}
      <div className={cn(
        "absolute bottom-8 text-[10px] font-serif opacity-30 transition-colors duration-500",
        theme.colors.text,
        side === 'left' ? "left-8" : "right-8"
      )}>
        {side === 'left' ? "L" : "R"}
      </div>
    </div>
  );
};

const BookmarkTab: React.FC<{ 
  label: string; 
  icon: React.ReactNode;
  isActive: boolean; 
  onClick: () => void;
  color: string;
  activeColor: string;
  theme: JournalTheme;
}> = ({ label, icon, isActive, onClick, color, activeColor, theme }) => {
  return (
    <motion.button
      onClick={onClick}
      className={cn(
        "w-12 h-14 flex items-center justify-center rounded-r-lg shadow-sm border transition-all relative overflow-hidden",
        theme.colors.border,
        isActive ? activeColor + " -right-4 w-16" : color + " -right-2 hover:-right-3"
      )}
      whileTap={{ scale: 0.95 }}
    >
      <div className={cn(
        "transition-colors",
        isActive ? "text-white scale-110" : theme.colors.textMuted
      )}>
        {icon}
      </div>
      {isActive && (
        <div className="absolute right-6 rotate-90 whitespace-nowrap text-[8px] font-bold uppercase tracking-[0.2em] text-white/40">
          {label}
        </div>
      )}
    </motion.button>
  );
};

export default JournalBook;
