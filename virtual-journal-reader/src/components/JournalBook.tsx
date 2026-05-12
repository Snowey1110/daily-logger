import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { cn, JournalEntry, JournalSection } from '../lib/utils';
import Cover from './Cover';
import Navigation from './Navigation';
import { DrawingCanvas } from './DrawingCanvas';
import { ChevronLeft, ChevronRight, MessageSquare, BrainCircuit, Type } from 'lucide-react';

export type JournalAction = 'sketch' | 'edit';

interface PageContent {
  type: 'text' | 'sketch' | 'empty';
  content?: string;
  /** 1-based page within journal text stream for this column layout */
  journalPage?: number;
  /** 1-based page within STT / AI column (split layout) */
  secondaryPage?: number;
}

interface Spread {
  entryId: string;
  date: string;
  time: string;
  left: PageContent;
  right: PageContent;
  isFirstSpread: boolean;
}

const MAX_CHARS_PER_PAGE = 800;

const splitTextIntoPages = (text: string): string[] => {
  if (!text) return [''];
  const pages: string[] = [];
  let currentText = text;

  while (currentText.length > 0) {
    if (currentText.length <= MAX_CHARS_PER_PAGE) {
      pages.push(currentText);
      break;
    }

    let sliceIndex = currentText.lastIndexOf(' ', MAX_CHARS_PER_PAGE);
    if (sliceIndex === -1) sliceIndex = MAX_CHARS_PER_PAGE;

    pages.push(currentText.substring(0, sliceIndex).trim());
    currentText = currentText.substring(sliceIndex).trim();
  }

  return pages;
};

function buildJournalBookmarkSpreads(entry: JournalEntry): Spread[] {
  const pages = splitTextIntoPages(entry.journal);
  const spreads: Spread[] = [];
  let sketchPlaced = false;
  let journalPage = 1;
  let i = 0;

  while (i < pages.length) {
    const isFirst = spreads.length === 0;
    const leftText = pages[i];
    const rightText = pages[i + 1];

    if (rightText !== undefined) {
      spreads.push({
        entryId: entry.id,
        date: entry.date,
        time: entry.time,
        left: { type: 'text', content: leftText, journalPage },
        right: { type: 'text', content: rightText, journalPage: journalPage + 1 },
        isFirstSpread: isFirst,
      });
      journalPage += 2;
      i += 2;
    } else {
      let right: PageContent = { type: 'empty' };
      if (entry.sketch) {
        right = { type: 'sketch', content: entry.sketch };
        sketchPlaced = true;
      }
      spreads.push({
        entryId: entry.id,
        date: entry.date,
        time: entry.time,
        left: { type: 'text', content: leftText, journalPage },
        right,
        isFirstSpread: isFirst,
      });
      journalPage += 1;
      i += 1;
    }
  }

  if (entry.sketch && !sketchPlaced) {
    spreads.push({
      entryId: entry.id,
      date: entry.date,
      time: entry.time,
      left: { type: 'empty' },
      right: { type: 'sketch', content: entry.sketch },
      isFirstSpread: false,
    });
  }

  return spreads;
}

function buildSplitBookmarkSpreads(entry: JournalEntry, section: 'stt' | 'ai'): Spread[] {
  const jp = splitTextIntoPages(entry.journal);
  const otherRaw = section === 'stt' ? entry.speechToText : entry.aiReport;
  const op = splitTextIntoPages(otherRaw);
  const n = Math.max(jp.length, op.length, 1);
  const spreads: Spread[] = [];
  let sketchPlaced = false;

  for (let idx = 0; idx < n; idx++) {
    const isFirst = idx === 0;
    spreads.push({
      entryId: entry.id,
      date: entry.date,
      time: entry.time,
      left: { type: 'text', content: jp[idx] ?? '', journalPage: idx + 1 },
      right: { type: 'text', content: op[idx] ?? '', secondaryPage: idx + 1 },
      isFirstSpread: isFirst,
    });
  }

  if (entry.sketch) {
    const last = spreads[spreads.length - 1];
    const rightEmpty =
      last.right.type === 'empty' ||
      (last.right.type === 'text' && !(last.right.content || '').trim());
    if (rightEmpty) {
      last.right = { type: 'sketch', content: entry.sketch };
      sketchPlaced = true;
    }
  }

  if (entry.sketch && !sketchPlaced) {
    spreads.push({
      entryId: entry.id,
      date: entry.date,
      time: entry.time,
      left: { type: 'empty' },
      right: { type: 'sketch', content: entry.sketch },
      isFirstSpread: false,
    });
  }

  return spreads;
}

function buildSpreadsForEntry(entry: JournalEntry, section: JournalSection): Spread[] {
  if (section === 'journal') {
    return buildJournalBookmarkSpreads(entry);
  }
  return buildSplitBookmarkSpreads(entry, section);
}

async function fetchEntries(): Promise<{
  entries: JournalEntry[];
  error?: string;
  appName: string;
}> {
  const res = await fetch('/api/entries');
  const data = await res.json();
  const appName =
    typeof data.appName === 'string' && data.appName.trim() ? data.appName.trim() : 'Daily Logger';
  return {
    entries: Array.isArray(data.entries) ? data.entries : [],
    error: typeof data.error === 'string' ? data.error : undefined,
    appName,
  };
}

const JournalBook: React.FC = () => {
  const [entries, setEntries] = useState<JournalEntry[]>([]);
  const [appTitle, setAppTitle] = useState('Daily Logger');
  const [loadError, setLoadError] = useState<string | null>(null);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [currentPage, setCurrentPage] = useState(0);
  const [activeSection, setActiveSection] = useState<JournalSection>('journal');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('asc');
  const [isSketching, setIsSketching] = useState(false);
  const [editingEntryId, setEditingEntryId] = useState<string | null>(null);
  const [editingSection, setEditingSection] = useState<JournalSection | null>(null);
  const [editedText, setEditedText] = useState('');

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const { entries: rows, error, appName } = await fetchEntries();
        if (cancelled) return;
        setEntries(rows);
        setAppTitle(appName);
        setLoadError(error ?? null);
      } catch {
        if (!cancelled) {
          setLoadError('Could not load journal data.');
          setEntries([]);
        }
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    document.title = `${appTitle} — Virtual Journal Reader`;
  }, [appTitle]);

  const sortedEntries = useMemo(() => {
    const list = [...entries];
    list.sort((a, b) => {
      const ai = `${a.isoDate ?? a.date}|${String(a.rowIndex ?? 0).padStart(6, '0')}`;
      const bi = `${b.isoDate ?? b.date}|${String(b.rowIndex ?? 0).padStart(6, '0')}`;
      const cmp = ai.localeCompare(bi);
      return sortOrder === 'asc' ? cmp : -cmp;
    });
    return list;
  }, [entries, sortOrder]);

  const getSpreads = useCallback(() => {
    const allSpreads: Spread[] = [];
    sortedEntries.forEach((entry) => {
      allSpreads.push(...buildSpreadsForEntry(entry, activeSection));
    });
    return allSpreads;
  }, [sortedEntries, activeSection]);

  const spreads = getSpreads();
  const totalPages = spreads.length + 1;

  const handleNext = useCallback(() => {
    if (currentPage < totalPages - 1) setCurrentPage((prev) => prev + 1);
  }, [currentPage, totalPages]);

  const handlePrev = useCallback(() => {
    if (currentPage > 0) setCurrentPage((prev) => prev - 1);
  }, [currentPage]);

  const handleAction = (action: JournalAction) => {
    if (currentPage === 0) return;
    const spread = spreads[currentPage - 1];
    if (!spread) return;

    if (action === 'sketch') {
      setIsSketching(true);
    } else {
      setSaveError(null);
      setEditingEntryId(spread.entryId);
      setEditingSection(activeSection);
      const entry = entries.find((e) => e.id === spread.entryId);
      if (!entry) return;
      if (activeSection === 'journal') setEditedText(entry.journal);
      else if (activeSection === 'stt') setEditedText(entry.speechToText);
      else setEditedText(entry.aiReport);
    }
  };

  const handleSaveText = async () => {
    if (!editingEntryId || !editingSection) return;
    setSaveError(null);
    const body: Record<string, string> = { id: editingEntryId };
    if (editingSection === 'journal') body.journal = editedText;
    if (editingSection === 'stt') body.speechToText = editedText;
    if (editingSection === 'ai') body.aiReport = editedText;
    try {
      const res = await fetch('/api/entry', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      const data = await res.json();
      if (!data.ok) {
        setSaveError(data.error || 'Save failed.');
        return;
      }
      const { entries: rows, error, appName } = await fetchEntries();
      setEntries(rows);
      setAppTitle(appName);
      setLoadError(error ?? null);
      setEditingEntryId(null);
      setEditingSection(null);
    } catch {
      setSaveError('Network error while saving.');
    }
  };

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (editingEntryId) return;
      const key = e.key.toLowerCase();
      if (key === 'd' || key === 'arrowright') handleNext();
      if (key === 'a' || key === 'arrowleft') handlePrev();
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleNext, handlePrev, editingEntryId]);

  const handleDateSelect = (date: Date) => {
    const dateStr = date.toLocaleDateString('en-US', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric',
    });
    const index = spreads.findIndex((s) => s.date === dateStr && s.isFirstSpread);
    if (index !== -1) {
      setCurrentPage(index + 1);
    }
  };

  const saveSketch = async (dataUrl: string) => {
    if (currentPage > 0) {
      const spread = spreads[currentPage - 1];
      try {
        const res = await fetch('/api/sketch', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ id: spread.entryId, dataUrl }),
        });
        const data = await res.json();
        if (!data.ok) {
          setSaveError(data.error || 'Could not save sketch.');
          setIsSketching(false);
          return;
        }
        const { entries: rows, error, appName } = await fetchEntries();
        setEntries(rows);
        setAppTitle(appName);
        setLoadError(error ?? null);
      } catch {
        setSaveError('Network error while saving sketch.');
      }
    }
    setIsSketching(false);
  };

  const getPageContent = () => spreads[currentPage - 1];

  const isEditing = !!editingEntryId && !!editingSection;

  return (
    <div className="flex h-dvh max-h-dvh min-h-0 flex-col overflow-hidden bg-[#2c1e14] text-[#d9c5b2] select-none font-sans">
      <div className="absolute top-6 left-8 text-[#d9c5b2] opacity-80 z-10">
        <h1 className="text-2xl tracking-widest uppercase font-light font-serif">
          {appTitle}{' '}
          <span className="text-xs opacity-50 block tracking-normal font-sans">Virtual Journal Reader</span>
        </h1>
      </div>

      {(loadError || saveError) && (
        <div className="absolute top-20 left-8 right-8 z-50 max-w-xl rounded-lg border border-amber-500/40 bg-black/50 px-4 py-2 text-sm text-amber-100 font-sans">
          {loadError && <p>{loadError}</p>}
          {saveError && <p>{saveError}</p>}
        </div>
      )}

      <Navigation
        currentPage={currentPage}
        totalPages={totalPages}
        onPageJump={setCurrentPage}
        onPrev={handlePrev}
        onNext={handleNext}
        onAction={handleAction}
        onToggleSort={() => setSortOrder((prev) => (prev === 'asc' ? 'desc' : 'asc'))}
        sortOrder={sortOrder}
        onDateSelect={handleDateSelect}
        isEditTextOpen={isEditing}
        onSaveText={() => void handleSaveText()}
        activeSection={activeSection}
      />

      <div className="flex-1 flex min-h-0 items-center justify-center p-3 md:p-5 perspective-[2000px]">
        <div className="relative mx-auto flex aspect-[1.4/1] h-auto max-h-full min-h-0 w-[min(100%,72rem,92vw)] max-w-full shrink shadow-[0_50px_100px_-20px_rgba(0,0,0,0.5)] rounded-xl">
          <div className="absolute inset-0 bg-black/40 -z-10 rounded-xl translate-x-2 translate-y-2 blur-2xl" />

          <div
            className={cn(
              'absolute left-1/2 -ml-0.5 top-0 bottom-0 w-px bg-black/10 z-20 shadow-[0_0_10px_rgba(0,0,0,0.1)]',
              currentPage === 0 && 'hidden',
            )}
          />

          <AnimatePresence mode="wait">
            {currentPage === 0 ? (
              <motion.div
                key="cover"
                className="h-full min-h-0 w-full"
                initial={{ rotateY: -10 }}
                animate={{ rotateY: 0 }}
                exit={{ rotateY: -90, opacity: 0 }}
                transition={{ duration: 0.8, ease: 'easeInOut' }}
                style={{ transformOrigin: 'left center' }}
              >
                <Cover
                  title={appTitle}
                  onClick={() => {
                    setCurrentPage(1);
                  }}
                />
              </motion.div>
            ) : (
              <motion.div
                key={`page-pair-${currentPage}`}
                className="relative flex h-full min-h-0 w-full overflow-hidden rounded-lg bg-[#fdfaf2] shadow-2xl"
                initial={{ scale: 0.95, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                transition={{ duration: 0.5 }}
              >
                <div className="absolute inset-0 opacity-10 pointer-events-none bg-[url('https://www.transparenttextures.com/patterns/natural-paper.png')]" />

                {!editingEntryId && (
                  <>
                    <button
                      type="button"
                      aria-label="Previous page"
                      className="absolute left-0 top-0 bottom-0 w-[12%] z-40 cursor-pointer opacity-0 hover:opacity-100 hover:bg-black/[0.03]"
                      onClick={handlePrev}
                    />
                    <button
                      type="button"
                      aria-label="Next page"
                      className="absolute right-0 top-0 bottom-0 w-[12%] z-40 cursor-pointer opacity-0 hover:opacity-100 hover:bg-black/[0.03]"
                      onClick={handleNext}
                    />
                  </>
                )}

                <Page
                  spread={getPageContent()}
                  activeSection={activeSection}
                  side="left"
                  editingEntryId={editingEntryId}
                  editingSection={editingSection}
                  editText={editedText}
                  onTextChange={setEditedText}
                />
                <Page
                  spread={getPageContent()}
                  activeSection={activeSection}
                  side="right"
                  editingEntryId={editingEntryId}
                  editingSection={editingSection}
                  editText={editedText}
                  onTextChange={setEditedText}
                />
              </motion.div>
            )}
          </AnimatePresence>

          {currentPage > 0 && (
            <div className="absolute left-full ml-1 top-20 flex flex-col space-y-1 z-30">
              <BookmarkTab
                label="JOURNAL"
                tabNumber="01"
                icon={<Type size={16} />}
                isActive={activeSection === 'journal'}
                onClick={() => setActiveSection('journal')}
              />
              <BookmarkTab
                label="SPEECH"
                tabNumber="02"
                icon={<MessageSquare size={16} />}
                isActive={activeSection === 'stt'}
                onClick={() => setActiveSection('stt')}
              />
              <BookmarkTab
                label="AI REPORT"
                tabNumber="03"
                icon={<BrainCircuit size={16} />}
                isActive={activeSection === 'ai'}
                onClick={() => setActiveSection('ai')}
              />
            </div>
          )}
        </div>
      </div>

      {isSketching && (
        <DrawingCanvas
          onSave={saveSketch}
          onClose={() => setIsSketching(false)}
          initialData={
            currentPage > 0 && spreads[currentPage - 1]
              ? entries.find((e) => e.id === spreads[currentPage - 1]!.entryId)?.sketch
              : undefined
          }
        />
      )}

      <footer className="shrink-0 flex flex-col items-center gap-2 px-4 pb-4 pt-2 text-center">
        <div className="flex items-center justify-center space-x-12 text-[#d9c5b2]/40">
          <button
            type="button"
            onClick={handlePrev}
            className="flex items-center space-x-2 group hover:text-[#d9c5b2] transition-colors"
          >
            <ChevronLeft className="w-5 h-5 group-hover:-translate-x-1 transition-transform" />
            <span className="text-sm uppercase tracking-widest font-sans">Previous</span>
          </button>
          <div className="flex items-center space-x-3">
            <div
              className={cn('w-1 h-1 rounded-full', currentPage < 2 ? 'bg-white/50' : 'bg-white/20')}
            />
            <div
              className={cn(
                'w-1.5 h-1.5 rounded-full',
                currentPage >= 2 && currentPage < totalPages ? 'bg-white/50' : 'bg-white/20',
              )}
            />
            <div
              className={cn(
                'w-1 h-1 rounded-full',
                currentPage >= totalPages - 1 ? 'bg-white/50' : 'bg-white/20',
              )}
            />
          </div>
          <button
            type="button"
            onClick={handleNext}
            className="flex items-center space-x-2 group hover:text-[#d9c5b2] transition-colors"
          >
            <span className="text-sm uppercase tracking-widest font-sans">Next</span>
            <ChevronRight className="w-5 h-5 group-hover:translate-x-1 transition-transform" />
          </button>
        </div>
        <div className="text-[10px] uppercase tracking-[0.3em] text-[#d9c5b2]/20 font-sans leading-relaxed">
          Use <kbd className="bg-white/5 px-1 rounded mx-1">A / D</kbd> or{' '}
          <kbd className="bg-white/5 px-1 rounded mx-1">Arrows</kbd> to flip pages
        </div>
      </footer>
    </div>
  );
};

const Page: React.FC<{
  spread?: Spread;
  activeSection: JournalSection;
  side: 'left' | 'right';
  editingEntryId: string | null;
  editingSection: JournalSection | null;
  editText?: string;
  onTextChange?: (text: string) => void;
}> = ({ spread, activeSection, side, editingEntryId, editingSection, editText, onTextChange }) => {
  if (!spread) {
    return (
      <div
        className={cn(
          'flex-1 h-full min-h-0 p-8 pr-10 flex items-center justify-center font-serif',
          side === 'right' && 'bg-[#fbf8ef] pl-10 pr-8',
        )}
      >
        <p className="text-zinc-400 italic tracking-widest uppercase text-xs opacity-40">
          The rest is still unwritten
        </p>
      </div>
    );
  }

  const pContent = side === 'left' ? spread.left : spread.right;

  if (pContent.type === 'sketch') {
    return (
      <div className="flex-1 h-full min-h-0 bg-[#fbf8ef] relative group overflow-hidden">
        <img
          src={pContent.content}
          alt="Full page sketch"
          className="w-full h-full object-contain opacity-90 mix-blend-multiply group-hover:opacity-100 transition-opacity bg-[#fbf8ef]"
        />
        <div className="absolute inset-0 border-l border-black/5 pointer-events-none" />
        <div className="absolute bottom-12 right-12 text-black/10 text-[9px] font-sans font-bold uppercase tracking-[0.3em]">
          {spread.date} · Digital Sketch
        </div>
      </div>
    );
  }

  if (pContent.type === 'empty') {
    return (
      <div
        className={cn(
          'flex-1 h-full min-h-0 p-8 pr-10 flex items-center justify-center font-serif',
          side === 'right' && 'bg-[#fbf8ef] pl-10 pr-8 border-l border-black/5',
        )}
      >
        <p className="text-zinc-400 italic tracking-widest uppercase text-[10px] opacity-20">
          (This space intentionally left blank)
        </p>
      </div>
    );
  }

  const leftTitle = () => {
    if (activeSection !== 'journal' && side === 'left') return 'Journal';
    if (activeSection === 'journal') return 'Daily Reflection';
    return 'Journal';
  };

  const rightTitle = () => {
    switch (activeSection) {
      case 'stt':
        return 'Voice Transcript';
      case 'ai':
        return 'Intelligence Analysis';
      default:
        return 'Daily Reflection';
    }
  };

  const sectionTitle = side === 'left' ? leftTitle() : rightTitle();
  const pageLabel =
    side === 'left'
      ? pContent.journalPage
      : activeSection === 'journal'
        ? pContent.journalPage
        : pContent.secondaryPage;

  const showBigHeader = spread.isFirstSpread && side === 'left';
  const showSubHeader =
    !showBigHeader &&
    pContent.type === 'text' &&
    (pageLabel !== undefined ? pageLabel >= 1 : false);

  const showEditor =
    editingEntryId === spread.entryId &&
    editingSection &&
    ((editingSection === 'journal' &&
      activeSection === 'journal' &&
      side === 'left' &&
      spread.isFirstSpread &&
      pContent.journalPage === 1) ||
      (editingSection === 'stt' &&
        activeSection === 'stt' &&
        side === 'right' &&
        spread.isFirstSpread &&
        pContent.secondaryPage === 1) ||
      (editingSection === 'ai' &&
        activeSection === 'ai' &&
        side === 'right' &&
        spread.isFirstSpread &&
        pContent.secondaryPage === 1));

  return (
    <div
      className={cn(
        'flex-1 h-full min-h-0 p-8 pr-10 flex flex-col relative overflow-hidden font-serif',
        side === 'right' ? 'pl-10 pr-8 bg-[#fbf8ef]' : 'bg-[#fdfaf2] border-r border-black/5',
      )}
    >
      {showBigHeader && (
        <div className="border-b border-black/5 pb-4 mb-6">
          <span className="text-xs uppercase tracking-widest text-black/40 font-sans">
            {spread.date} · {spread.time}
          </span>
          <h2 className="text-2xl font-light text-slate-800 leading-tight mt-1 font-serif">
            Journal entry
          </h2>
        </div>
      )}

      {showSubHeader && (
        <div className="border-b border-black/10 border-dashed pb-2 mb-6 flex justify-between items-end">
          <h3 className="text-sm font-light text-slate-400 font-serif italic tracking-wide">
            {pageLabel !== undefined && pageLabel > 1 ? `${sectionTitle} (cont.)` : sectionTitle}
          </h3>
          <span className="text-[10px] text-black/20 font-sans">{spread.date}</span>
        </div>
      )}

      <div className="flex-1 text-slate-700 leading-relaxed text-base space-y-3 overflow-y-auto pr-2">
        {showEditor ? (
          <textarea
            value={editText}
            onChange={(e) => onTextChange?.(e.target.value)}
            className="w-full min-h-[60%] bg-transparent border border-black/10 rounded-md p-2 focus:ring-1 focus:ring-amber-700/30 resize-y font-serif text-slate-800 leading-relaxed"
            placeholder="Edit text…"
            autoFocus
          />
        ) : (
          (pContent.content || '').split('\n').map((para, i) => <p key={i}>{para}</p>)
        )}
      </div>

      {side === 'right' && (
        <div className="mt-auto pt-6 border-t border-black/5 text-right">
          <p className="text-sm text-slate-500 italic font-serif">
            “The machine sees what we often overlook…”
          </p>
        </div>
      )}

      <div className="mt-8 flex justify-center shrink-0">
        <div className="w-16 h-1 bg-black/5 rounded-full" />
      </div>
    </div>
  );
};

const BookmarkTab: React.FC<{
  label: string;
  tabNumber: string;
  icon: React.ReactNode;
  isActive: boolean;
  onClick: () => void;
}> = ({ label, tabNumber, icon, isActive, onClick }) => {
  return (
    <motion.button
      type="button"
      onClick={onClick}
      className={cn(
        'px-6 py-3 rounded-r-md shadow-sm border-l-4 flex flex-col cursor-pointer transition-all w-32 items-start font-sans',
        isActive
          ? 'bg-[#e8dccb] text-[#5c4a36] border-[#8c7a66] translate-x-2'
          : 'bg-[#f4ead5]/80 text-[#8c7a66] border-[#8c7a66]/20 hover:bg-[#e8dccb] hover:translate-x-1',
      )}
      whileTap={{ scale: 0.98 }}
    >
      <span className="text-[10px] font-bold uppercase tracking-tighter opacity-60">{tabNumber}</span>
      <span className="text-[11px] font-bold tracking-tight uppercase flex items-center gap-2">
        {icon}
        {label}
      </span>
    </motion.button>
  );
};

export default JournalBook;
