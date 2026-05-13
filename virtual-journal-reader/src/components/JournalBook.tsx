import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { cn, JournalEntry, JournalSection, PositionedSketch, PageImage, PageOverlay } from '../lib/utils';
import Cover from './Cover';
import Navigation from './Navigation';
import { DrawingCanvas } from './DrawingCanvas';
import { SketchPlacer } from './SketchPlacer';
import { CompositePageEditor } from './CompositePageEditor';
import { ChevronLeft, ChevronRight, Type, MessageSquare, BrainCircuit } from 'lucide-react';
import { useReaderT } from '../readerI18n';
import { useTheme } from './ThemeProvider';
import type { JournalTheme } from '../types/theme';

export type JournalAction = 'sketch' | 'edit';

interface PageContent {
  type: 'text' | 'sketch' | 'empty';
  content?: string;
  journalPage?: number;
  secondaryPage?: number;
  sourceEntryId?: string;
  sourceDate?: string;
  sourceTime?: string;
  showJournalEntryHeader?: boolean;
  /** When journal text is empty, indicates which section was used as fallback on the left page */
  leftFallbackSection?: 'stt' | 'ai';
}

interface Spread {
  entryId: string;
  date: string;
  time: string;
  left: PageContent;
  right: PageContent;
  isFirstSpread: boolean;
}

/* ──── localStorage helpers ──── */

const READER_SORT_KEY = 'virtualJournalReader.sortOrder';

function readStored<T extends string>(key: string, allowed: T[], fallback: T): T {
  try {
    const v = localStorage.getItem(key);
    if (v && (allowed as string[]).includes(v)) return v as T;
  } catch { /* ignore */ }
  return fallback;
}

function persistStored(key: string, value: string): void {
  try { localStorage.setItem(key, value); } catch { /* ignore */ }
}

/** Convert "YYYY-MM-DD" → "MM/DD/YYYY" for display when the Date cell is empty. */
function isoToDisplayDate(iso: string): string {
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? `${m[2]}/${m[3]}/${m[1]}` : iso;
}

/* ──── helpers for building left-page content ──── */

function buildLeftPage(
  entry: JournalEntry,
  displayDate: string,
  displayTime: string,
): { left: PageContent; leftFallback: 'journal' | 'stt' | 'ai' } {
  const journalText = (entry.journal || '').trim();
  const sttText = (entry.speechToText || '').trim();
  const aiText = (entry.aiReport || '').trim();

  let leftFallback: 'journal' | 'stt' | 'ai' = 'journal';
  let leftText = journalText;
  if (!leftText && sttText) {
    leftFallback = 'stt';
    leftText = sttText;
  } else if (!leftText && aiText) {
    leftFallback = 'ai';
    leftText = aiText;
  }

  const left: PageContent = {
    type: 'text',
    content: leftText,
    journalPage: 1,
    sourceEntryId: entry.id,
    sourceDate: displayDate,
    sourceTime: displayTime,
    showJournalEntryHeader: true,
    leftFallbackSection: leftFallback !== 'journal' ? leftFallback : undefined,
  };

  return { left, leftFallback };
}

function entryDisplayDate(entry: JournalEntry): string {
  return entry.date || (entry.isoDate ? isoToDisplayDate(entry.isoDate) : '');
}

/* ──── page-break helper: split on literal \n (backslash-n) ──── */

function splitByPageBreak(text: string): string[] {
  if (!text) return [''];
  return text.split('\\n');
}

/** In non-journal modes, treat literal \n as a regular newline instead of a page break. */
function pageBreakToNewline(text: string): string {
  return text.replaceAll('\\n', '\n');
}

function applyPageBreaks(rawSpreads: Spread[]): Spread[] {
  const result: Spread[] = [];

  for (const spread of rawSpreads) {
    const { left, right } = spread;

    const segments: PageContent[] = [];

    if (left.type === 'text' && left.content && left.content.includes('\\n')) {
      const parts = splitByPageBreak(left.content);
      for (let i = 0; i < parts.length; i++) {
        segments.push({
          ...left,
          content: parts[i],
          showJournalEntryHeader: i === 0 ? left.showJournalEntryHeader : false,
          journalPage: i === 0 ? left.journalPage : undefined,
        });
      }
    } else {
      segments.push(left);
    }

    if (right.type === 'text' && right.content && right.content.includes('\\n')) {
      const parts = splitByPageBreak(right.content);
      for (const part of parts) {
        segments.push({ ...right, content: part });
      }
    } else {
      segments.push(right);
    }

    if (segments.length === 2) {
      result.push(spread);
      continue;
    }

    for (let i = 0; i < segments.length; i += 2) {
      const newLeft = segments[i];
      const newRight = i + 1 < segments.length ? segments[i + 1] : { type: 'empty' as const };

      result.push({
        entryId: spread.entryId,
        date: spread.date,
        time: spread.time,
        left: newLeft,
        right: newRight,
        isFirstSpread: i === 0 ? spread.isFirstSpread : false,
      });
    }
  }

  return result;
}

/* ──── unified spread builder (one spread per entry, scrollable) ──── */

interface PageItem {
  content: PageContent;
  entryId: string;
  date: string;
  time: string;
  sketchContents?: PageContent[];
}

function expandEntryToPages(
  entry: JournalEntry,
  displayDate: string,
  displayTime: string,
  left: PageContent,
  sketchPCs: PageContent[],
): PageItem[] {
  const textPages: PageContent[] = [];

  if (left.type === 'text' && left.content && left.content.includes('\\n')) {
    const parts = splitByPageBreak(left.content);
    for (let j = 0; j < parts.length; j++) {
      textPages.push({
        ...left,
        content: parts[j],
        showJournalEntryHeader: j === 0 ? left.showJournalEntryHeader : false,
        journalPage: j === 0 ? left.journalPage : undefined,
      });
    }
  } else {
    textPages.push(left);
  }

  return textPages.map((tp, j) => ({
    content: tp,
    entryId: entry.id,
    date: displayDate,
    time: displayTime,
    sketchContents: j === textPages.length - 1 && sketchPCs.length > 0 ? sketchPCs : undefined,
  }));
}

function buildUnifiedSpreads(
  sortedEntries: JournalEntry[],
  sketches: PositionedSketch[],
  activeSection: JournalSection,
): Spread[] {
  // Build a map from entryId -> ordered list of sketches
  const sketchesByEntry = new Map<string, PositionedSketch[]>();
  for (const sk of sketches) {
    const list = sketchesByEntry.get(sk.afterEntryId) ?? [];
    list.push(sk);
    sketchesByEntry.set(sk.afterEntryId, list);
  }

  const spreads: Spread[] = [];

  if (activeSection === 'journal') {
    // Expand all entries into page items (splitting by \n), then pair.
    const items: PageItem[] = [];

    for (const entry of sortedEntries) {
      const entrySketches = sketchesByEntry.get(entry.id) ?? [];
      const displayDate = entryDisplayDate(entry);
      const displayTime = entry.time;
      const { left } = buildLeftPage(entry, displayDate, displayTime);

      const sketchPCs: PageContent[] = entrySketches.map((sk) => ({
        type: 'sketch' as const,
        content: sk.dataUrl,
        sourceEntryId: entry.id,
        sourceDate: displayDate,
        sourceTime: displayTime,
      }));

      items.push(...expandEntryToPages(entry, displayDate, displayTime, left, sketchPCs));
    }

    // Pair page items into spreads
    let i = 0;
    while (i < items.length) {
      const item = items[i];

      if (item.sketchContents && item.sketchContents.length > 0) {
        // First sketch goes on the right of the text page
        spreads.push({
          entryId: item.entryId,
          date: item.date,
          time: item.time,
          left: item.content,
          right: item.sketchContents[0],
          isFirstSpread: true,
        });
        // Remaining sketches: pair them two at a time (left + right)
        let s = 1;
        while (s < item.sketchContents.length) {
          const leftSketch = item.sketchContents[s];
          const rightSketch = s + 1 < item.sketchContents.length ? item.sketchContents[s + 1] : null;
          spreads.push({
            entryId: item.entryId,
            date: item.date,
            time: item.time,
            left: leftSketch,
            right: rightSketch ?? { type: 'empty' },
            isFirstSpread: false,
          });
          s += rightSketch ? 2 : 1;
        }
        i += 1;
      } else {
        // Pair with the next page item
        const nextItem = i + 1 < items.length ? items[i + 1] : null;
        spreads.push({
          entryId: item.entryId,
          date: item.date,
          time: item.time,
          left: item.content,
          right: nextItem ? nextItem.content : { type: 'empty' },
          isFirstSpread: true,
        });
        i += nextItem ? 2 : 1;
      }
    }
  } else {
    // STT or AI mode: left page cycles through journal text then sketches,
    // right page always shows the secondary content (STT or AI report).
    for (let i = 0; i < sortedEntries.length; i++) {
      const entry = sortedEntries[i];
      const entrySketches = sketchesByEntry.get(entry.id) ?? [];
      const displayDate = entryDisplayDate(entry);
      const displayTime = entry.time;
      const { left, leftFallback } = buildLeftPage(entry, displayDate, displayTime);

      const leftResolved: PageContent = left.type === 'text' && left.content
        ? { ...left, content: pageBreakToNewline(left.content) }
        : left;

      const secondaryField = activeSection === 'stt' ? 'speechToText' : 'aiReport';
      let secondaryText = (entry[secondaryField] || '').trim();
      if (leftFallback === activeSection) {
        secondaryText = activeSection === 'stt'
          ? (entry.aiReport || '').trim()
          : (entry.speechToText || '').trim();
      }
      secondaryText = pageBreakToNewline(secondaryText);

      const secondaryPC: PageContent = { type: 'text', content: secondaryText, secondaryPage: 1, sourceEntryId: entry.id, sourceDate: displayDate, sourceTime: displayTime };

      // First spread: journal text on left, secondary on right
      spreads.push({
        entryId: entry.id,
        date: displayDate,
        time: displayTime,
        left: leftResolved,
        right: secondaryPC,
        isFirstSpread: true,
      });

      // Each sketch gets its own spread with secondary content pinned on the right
      for (const sk of entrySketches) {
        const sketchPC: PageContent = {
          type: 'sketch',
          content: sk.dataUrl,
          sourceEntryId: entry.id,
          sourceDate: displayDate,
          sourceTime: displayTime,
        };
        spreads.push({
          entryId: entry.id,
          date: displayDate,
          time: displayTime,
          left: sketchPC,
          right: secondaryPC,
          isFirstSpread: false,
        });
      }
    }
  }

  return spreads;
}

/* ──── spread navigation helpers ──── */

function spreadPrimaryEntryId(s: Spread): string {
  if (s.left.type === 'text' && (s.left.content || '').trim()) return s.left.sourceEntryId ?? s.entryId;
  if (s.left.type === 'sketch') return s.left.sourceEntryId ?? s.entryId;
  if (s.right.type === 'text' && (s.right.content || '').trim()) return s.right.sourceEntryId ?? s.entryId;
  if (s.right.type === 'sketch') return s.right.sourceEntryId ?? s.entryId;
  return s.entryId;
}

function spreadTouchesEntryId(s: Spread, entryId: string): boolean {
  const own = (p: PageContent) => {
    if (p.type === 'text' && (p.content || '').trim()) return p.sourceEntryId ?? s.entryId;
    if (p.type === 'sketch') return p.sourceEntryId ?? s.entryId;
    return null;
  };
  return own(s.left) === entryId || own(s.right) === entryId;
}

function spreadHasCalendarHeader(s: Spread, dateStr: string): boolean {
  for (const p of [s.left, s.right]) {
    if (p.type === 'text' && p.showJournalEntryHeader && (p.sourceDate ?? '') === dateStr) return true;
  }
  return false;
}

function firstSpreadIndexForEntry(spreads: Spread[], entryId: string): number {
  const textIx = spreads.findIndex(
    (s) =>
      (s.left.type === 'text' && (s.left.sourceEntryId ?? s.entryId) === entryId && s.left.journalPage === 1 && s.left.showJournalEntryHeader) ||
      (s.right.type === 'text' && (s.right.sourceEntryId ?? s.entryId) === entryId && s.right.journalPage === 1 && s.right.showJournalEntryHeader),
  );
  if (textIx >= 0) return textIx;
  return spreads.findIndex((s) => spreadTouchesEntryId(s, entryId));
}

/* ──── data fetching ──── */

async function fetchData(): Promise<{
  entries: JournalEntry[];
  sketches: PositionedSketch[];
  overlays: Record<string, PageOverlay>;
  error?: string;
  appName: string;
}> {
  const res = await fetch('/api/entries');
  const data = await res.json();
  const appName = typeof data.appName === 'string' && data.appName.trim() ? data.appName.trim() : 'Daily Logger';
  const rawOverlays = (typeof data.overlays === 'object' && data.overlays !== null) ? data.overlays : {};
  const overlays: Record<string, PageOverlay> = {};
  for (const [k, v] of Object.entries(rawOverlays)) {
    const val = v as Record<string, unknown>;
    overlays[k] = {
      entryId: k,
      sketchDataUrl: (typeof val.sketchDataUrl === 'string' ? val.sketchDataUrl : undefined),
      images: Array.isArray(val.images) ? val.images as PageImage[] : [],
      layerOrder: Array.isArray(val.layerOrder) ? val.layerOrder as ('text' | 'sketch' | 'images')[] : ['text', 'sketch', 'images'],
    };
  }
  return {
    entries: Array.isArray(data.entries) ? data.entries : [],
    sketches: Array.isArray(data.sketches) ? data.sketches : [],
    overlays,
    error: typeof data.error === 'string' ? data.error : undefined,
    appName,
  };
}

/* ──── main component ──── */

const JournalBook: React.FC = () => {
  const { t } = useReaderT();
  const { coverTheme, bgTheme } = useTheme();
  const [entries, setEntries] = useState<JournalEntry[]>([]);
  const [sketches, setSketches] = useState<PositionedSketch[]>([]);
  const [overlays, setOverlays] = useState<Record<string, PageOverlay>>({});
  const [appTitle, setAppTitle] = useState('Daily Logger');
  const [loadError, setLoadError] = useState<string | null>(null);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [currentPage, setCurrentPage] = useState(0);
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>(() => readStored(READER_SORT_KEY, ['asc', 'desc'], 'asc'));
  const [activeSection, setActiveSection] = useState<JournalSection>('journal');

  const [showSketchPlacer, setShowSketchPlacer] = useState(false);
  const [editingSketchId, setEditingSketchId] = useState<string | null>(null);

  const [compositeEditorEntry, setCompositeEditorEntry] = useState<{ entryId: string; defaultLayer: 'text' | 'sketch' | 'images'; pageWidth: number; pageHeight: number } | null>(null);

  const bookSpreadRef = useRef<HTMLDivElement>(null);

  /* ── load data ── */

  const reload = useCallback(async () => {
    try {
      const { entries: rows, sketches: sk, overlays: ov, error, appName } = await fetchData();
      setEntries(rows);
      setSketches(sk);
      setOverlays(ov);
      setAppTitle(appName);
      setLoadError(error ?? null);
    } catch {
      setLoadError(t('errLoadData'));
      setEntries([]);
      setSketches([]);
      setOverlays({});
    }
  }, [t]);

  useEffect(() => { void reload(); }, [reload]);

  useEffect(() => {
    document.title = `${appTitle} — ${t('docTitleSuffix')}`;
  }, [appTitle, t]);

  /* ── sorted entries ── */

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

  /* ── spreads ── */

  const spreads = useMemo(
    () => buildUnifiedSpreads(sortedEntries, sketches, activeSection),
    [sortedEntries, sketches, activeSection],
  );

  const spreadCount = spreads.length;

  useEffect(() => {
    if (spreadCount > 0 && currentPage > spreadCount) setCurrentPage(spreadCount);
  }, [spreadCount, currentPage]);

  /* ── navigation ── */

  const goToEntry = useCallback((entryId: string) => {
    const spl = buildUnifiedSpreads(sortedEntries, sketches, activeSection);
    let ix = firstSpreadIndexForEntry(spl, entryId);
    if (ix < 0) ix = spl.findIndex((s) => spreadTouchesEntryId(s, entryId));
    setCurrentPage(ix >= 0 ? ix + 1 : 1);
  }, [sortedEntries, sketches, activeSection]);

  const handlePageJump = useCallback((page: number) => {
    if (!Number.isFinite(page)) return;
    const maxP = spreads.length;
    if (maxP < 1) { setCurrentPage(0); return; }
    setCurrentPage(Math.min(Math.max(1, Math.floor(page)), maxP));
  }, [spreads]);

  const handleNext = useCallback(() => {
    if (currentPage >= spreadCount) return;
    setCurrentPage((p) => p + 1);
  }, [currentPage, spreadCount]);

  const handlePrev = useCallback(() => {
    if (currentPage <= 0) return;
    setCurrentPage((p) => Math.max(0, p - 1));
  }, [currentPage]);

  const handleDateSelect = (date: Date) => {
    const dateStr = date.toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
    const index = spreads.findIndex(
      (s) => spreadHasCalendarHeader(s, dateStr) || (s.date === dateStr && s.isFirstSpread),
    );
    if (index !== -1) setCurrentPage(index + 1);
  };

  /* ── keyboard nav ── */

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (compositeEditorEntry) return;
      const key = e.key.toLowerCase();
      if (key === 'd' || key === 'arrowright') handleNext();
      if (key === 'a' || key === 'arrowleft') handlePrev();
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleNext, handlePrev, compositeEditorEntry]);

  /* ── actions ── */

  const measurePageDims = (): { pageWidth: number; pageHeight: number } => {
    const el = bookSpreadRef.current;
    if (el) {
      const rect = el.getBoundingClientRect();
      return { pageWidth: Math.round(rect.width / 2), pageHeight: Math.round(rect.height) };
    }
    return { pageWidth: 500, pageHeight: 650 };
  };

  const handleAction = (action: JournalAction) => {
    if (action === 'sketch') {
      setShowSketchPlacer(true);
    } else {
      if (currentPage === 0) return;
      const spread = spreads[currentPage - 1];
      if (!spread) return;
      const leftEid = spread.left.sourceEntryId ?? spread.entryId;
      if (leftEid) {
        setCompositeEditorEntry({ entryId: leftEid, defaultLayer: 'text', ...measurePageDims() });
      }
    }
  };

  /* ── composite editor ── */

  const handleOpenCompositeEditor = (entryId: string, defaultLayer: 'text' | 'sketch') => {
    setCompositeEditorEntry({ entryId, defaultLayer, ...measurePageDims() });
  };

  const handleSaveComposite = async (
    text: string,
    sketchDataUrl: string,
    images: PageImage[],
    layerOrder: ('text' | 'sketch' | 'images')[],
  ) => {
    if (!compositeEditorEntry) return;
    const { entryId } = compositeEditorEntry;
    setSaveError(null);
    try {
      const entryBody: Record<string, string> = { id: entryId, journal: text };
      const res = await fetch('/api/entry', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(entryBody),
      });
      const data = await res.json();
      if (!data.ok) { setSaveError(data.error || t('errSaveFailed')); return; }

      const overlayBody = { entryId, sketchDataUrl, images, layerOrder };
      const res2 = await fetch('/api/page-overlay', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(overlayBody),
      });
      const data2 = await res2.json();
      if (!data2.ok) { setSaveError(data2.error || t('errSaveFailed')); return; }

      setCompositeEditorEntry(null);
      await reload();
    } catch { setSaveError(t('errNetworkSave')); }
  };

  /* ── sketch CRUD (backward compat for existing standalone sketches) ── */

  const handleEditSketch = (sketchId: string) => {
    setEditingSketchId(sketchId);
    setShowSketchPlacer(false);
  };

  const handleDeleteSketch = async (sketchId: string) => {
    try {
      const res = await fetch('/api/sketch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: sketchId, delete: true }),
      });
      const data = await res.json();
      if (!data.ok) { setSaveError(data.error || t('errSketchSave')); return; }
      await reload();
    } catch { setSaveError(t('errNetworkSketch')); }
  };

  const handleSaveSketchCanvas = async (dataUrl: string, existingId?: string) => {
    setSaveError(null);
    if (!dataUrl || !dataUrl.startsWith('data:')) {
      if (existingId) await handleDeleteSketch(existingId);
      setEditingSketchId(null);
      return;
    }
    try {
      const body: Record<string, unknown> = { dataUrl };
      if (existingId) body.id = existingId;
      const res = await fetch('/api/sketch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      const data = await res.json();
      if (!data.ok) { setSaveError(data.error || t('errSketchSave')); }
      await reload();
    } catch { setSaveError(t('errNetworkSketch')); }
    setEditingSketchId(null);
  };

  const handleDeleteEntry = async (entryId: string) => {
    setSaveError(null);
    try {
      const res = await fetch('/api/entry/delete', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: entryId }),
      });
      const data = await res.json();
      if (!data.ok) { setSaveError(data.error || t('errDeleteEntry')); return; }
      await reload();
    } catch { setSaveError(t('errDeleteEntry')); }
  };

  const handleCreatePage = async (date: string, time: string, _afterEntryId: string): Promise<string | null> => {
    setSaveError(null);
    try {
      const res = await fetch('/api/entry/create', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ date, time }),
      });
      const data = await res.json();
      if (!data.ok) { setSaveError(data.error || t('errCreatePage')); return null; }
      await reload();
      if (data.id) {
        goToEntry(data.id);
        return data.id as string;
      }
    } catch { setSaveError(t('errCreatePage')); }
    return null;
  };

  const isDrawing = editingSketchId !== null;
  const drawingInitialData = editingSketchId ? sketches.find((s) => s.id === editingSketchId)?.dataUrl : undefined;

  const handleBookmarkClick = (section: JournalSection) => {
    setActiveSection(section);
  };

  const currentSpread = currentPage > 0 ? spreads[currentPage - 1] : undefined;

  return (
    <div className="flex h-dvh max-h-dvh min-h-0 flex-col overflow-hidden select-none font-sans transition-colors duration-500" style={{ backgroundColor: bgTheme.colors.bg }}>
      <div className="absolute top-6 left-8 opacity-80 z-10" style={{ color: bgTheme.cover.accentText }}>
        <h1 className="text-2xl tracking-widest uppercase font-light font-serif">
          {appTitle}{' '}
          <span className="text-xs opacity-50 block tracking-normal font-sans">{t('readerSubtitle')}</span>
        </h1>
      </div>

      {(loadError || saveError) && (
        <div className="fixed top-6 left-1/2 -translate-x-1/2 z-[200] max-w-xl rounded-lg border border-amber-500/40 bg-black/80 px-6 py-3 text-sm text-amber-100 font-sans shadow-2xl backdrop-blur-sm">
          {loadError && <p>{loadError}</p>}
          {saveError && <p>{saveError}</p>}
        </div>
      )}

      <Navigation
        currentPage={currentPage}
        totalPages={spreadCount}
        onPageJump={handlePageJump}
        onPrev={handlePrev}
        onNext={handleNext}
        onAction={handleAction}
        onToggleSort={() =>
          setSortOrder((prev) => {
            const next = prev === 'asc' ? 'desc' : 'asc';
            persistStored(READER_SORT_KEY, next);
            return next;
          })
        }
        sortOrder={sortOrder}
        onDateSelect={handleDateSelect}
        isEditTextOpen={false}
        onSaveText={() => {}}
      />

      <div className="flex-1 flex min-h-0 items-center justify-center p-3 md:p-5 perspective-[2000px]">
        <div className="relative mx-auto flex aspect-[1.4/1] h-auto max-h-full min-h-0 w-[min(100%,72rem,92vw)] max-w-full shrink shadow-[0_50px_100px_-20px_rgba(0,0,0,0.5)] rounded-xl">
          <div className="absolute inset-0 bg-black/40 -z-10 rounded-xl translate-x-2 translate-y-2 blur-2xl" />

          <div
            className={cn(
              'absolute left-1/2 -ml-0.5 top-0 bottom-0 w-px z-20 shadow-[0_0_10px_rgba(0,0,0,0.1)]',
              currentPage === 0 && 'hidden',
            )}
            style={{ backgroundColor: bgTheme.colors.spine }}
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
                  theme={coverTheme}
                  onClick={() => {
                    const first = sortedEntries[0];
                    if (first) goToEntry(first.id);
                    else setCurrentPage(1);
                  }}
                />
              </motion.div>
            ) : (
              <motion.div
                ref={bookSpreadRef}
                key={`page-pair-${currentPage}`}
                className="relative flex h-full min-h-0 w-full overflow-hidden rounded-lg shadow-2xl transition-colors duration-500"
                style={{ backgroundColor: bgTheme.colors.bookInner }}
                initial={{ scale: 0.95, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                transition={{ duration: 0.5 }}
              >
                <div className="absolute inset-0 opacity-10 pointer-events-none bg-[url('https://www.transparenttextures.com/patterns/natural-paper.png')]" />

                {/* Click zones for page flip */}
                <>
                  <button
                    type="button"
                    aria-label={t('ariaPrevPage')}
                    className="absolute left-0 top-0 bottom-0 w-[12%] z-40 cursor-pointer opacity-0 hover:opacity-100 hover:bg-black/[0.03]"
                    onClick={handlePrev}
                  />
                  <button
                    type="button"
                    aria-label={t('ariaNextPage')}
                    className="absolute right-0 top-0 bottom-0 w-[12%] z-40 cursor-pointer opacity-0 hover:opacity-100 hover:bg-black/[0.03]"
                    onClick={handleNext}
                  />
                </>

                <Page
                  spread={currentSpread}
                  side="left"
                  activeSection={activeSection}
                  theme={bgTheme}
                  overlay={currentSpread?.left.sourceEntryId ? overlays[currentSpread.left.sourceEntryId] : undefined}
                />
                <Page
                  spread={currentSpread}
                  side="right"
                  activeSection={activeSection}
                  theme={bgTheme}
                  overlay={currentSpread?.right.sourceEntryId ? overlays[currentSpread.right.sourceEntryId] : undefined}
                />
              </motion.div>
            )}
          </AnimatePresence>

          {/* Bookmark tabs on the right edge of the book */}
          {currentPage > 0 && (
            <div className="absolute left-full ml-1 top-20 flex flex-col space-y-1 z-30">
              <BookmarkTab
                label={t('tabJournal')}
                tabNumber="01"
                icon={<Type size={16} />}
                isActive={activeSection === 'journal'}
                onClick={() => handleBookmarkClick('journal')}
                colors={bgTheme.colors.tabs.journal}
              />
              <BookmarkTab
                label={t('tabSpeech')}
                tabNumber="02"
                icon={<MessageSquare size={16} />}
                isActive={activeSection === 'stt'}
                onClick={() => handleBookmarkClick('stt')}
                colors={bgTheme.colors.tabs.stt}
              />
              <BookmarkTab
                label={t('tabAi')}
                tabNumber="03"
                icon={<BrainCircuit size={16} />}
                isActive={activeSection === 'ai'}
                onClick={() => handleBookmarkClick('ai')}
                colors={bgTheme.colors.tabs.ai}
              />
            </div>
          )}
        </div>
      </div>

      {/* Insert manager modal */}
      {showSketchPlacer && (
        <SketchPlacer
          entries={sortedEntries}
          sketches={sketches}
          sortOrder={sortOrder}
          onCreatePage={handleCreatePage}
          onOpenCompositeEditor={handleOpenCompositeEditor}
          onEditSketch={handleEditSketch}
          onDeleteSketch={handleDeleteSketch}
          onDeleteEntry={handleDeleteEntry}
          onClose={() => setShowSketchPlacer(false)}
        />
      )}

      {/* Drawing canvas (backward compat for editing existing standalone sketches) */}
      {isDrawing && (
        <DrawingCanvas
          onSave={handleSaveSketchCanvas}
          onClose={() => setEditingSketchId(null)}
          initialData={drawingInitialData}
          sketchId={editingSketchId ?? undefined}
        />
      )}

      {/* Composite page editor */}
      {compositeEditorEntry && (() => {
        const entry = entries.find((e) => e.id === compositeEditorEntry.entryId);
        const ov = overlays[compositeEditorEntry.entryId];
        return (
          <CompositePageEditor
            entryId={compositeEditorEntry.entryId}
            entryDate={entry?.date ?? ''}
            entryTime={entry?.time ?? ''}
            pageWidth={compositeEditorEntry.pageWidth}
            pageHeight={compositeEditorEntry.pageHeight}
            initialText={entry?.journal ?? ''}
            initialSketchDataUrl={ov?.sketchDataUrl}
            initialImages={ov?.images ?? []}
            initialLayerOrder={ov?.layerOrder ?? ['text', 'sketch', 'images']}
            defaultLayer={compositeEditorEntry.defaultLayer}
            onSave={handleSaveComposite}
            onClose={() => setCompositeEditorEntry(null)}
          />
        );
      })()}

      <footer className="shrink-0 flex flex-col items-center gap-2 px-4 pb-4 pt-2 text-center" style={{ color: bgTheme.cover.accentText }}>
        <div className="flex items-center justify-center space-x-12 opacity-40">
          <button
            type="button"
            onClick={handlePrev}
            className="flex items-center space-x-2 group hover:opacity-100 transition-opacity"
          >
            <ChevronLeft className="w-5 h-5 group-hover:-translate-x-1 transition-transform" />
            <span className="text-sm uppercase tracking-widest font-sans">{t('footerPrev')}</span>
          </button>
          <div className="flex items-center space-x-3">
            <div className={cn('w-1 h-1 rounded-full', currentPage < 2 ? 'bg-current/50' : 'bg-current/20')} />
            <div className={cn('w-1.5 h-1.5 rounded-full', currentPage >= 2 && currentPage < spreadCount ? 'bg-current/50' : 'bg-current/20')} />
            <div className={cn('w-1 h-1 rounded-full', currentPage >= spreadCount ? 'bg-current/50' : 'bg-current/20')} />
          </div>
          <button
            type="button"
            onClick={handleNext}
            className="flex items-center space-x-2 group hover:opacity-100 transition-opacity"
          >
            <span className="text-sm uppercase tracking-widest font-sans">{t('footerNext')}</span>
            <ChevronRight className="w-5 h-5 group-hover:translate-x-1 transition-transform" />
          </button>
        </div>
        <div className="text-[10px] uppercase tracking-[0.3em] opacity-20 font-sans leading-relaxed">
          {t('footerFlipHint')}
        </div>
      </footer>
    </div>
  );
};

/* ──── Page sub-component ──── */

const Page: React.FC<{
  spread?: Spread;
  side: 'left' | 'right';
  activeSection: JournalSection;
  theme: JournalTheme;
  overlay?: PageOverlay;
}> = ({ spread, side, activeSection, theme, overlay }) => {
  const { t } = useReaderT();
  if (!spread) {
    return (
      <div
        className="w-1/2 h-full min-h-0 p-8 pr-10 flex items-center justify-center font-serif"
        style={side === 'right' ? { backgroundColor: theme.colors.bookInner, paddingLeft: '2.5rem', paddingRight: '2rem' } : undefined}
      >
        <p className="italic tracking-widest uppercase text-xs opacity-40" style={{ color: theme.colors.textMuted }}>{t('pageEmpty')}</p>
      </div>
    );
  }

  const pContent = side === 'left' ? spread.left : spread.right;

  /* ── sketch page (standalone) ── */
  if (pContent.type === 'sketch') {
    const capDate = pContent.sourceDate ?? spread.date;
    return (
      <div className="w-1/2 h-full min-h-0 relative group overflow-hidden flex flex-col" style={{ backgroundColor: theme.colors.bookInner }}>
        <div className="flex-1 min-h-0 relative">
          <img
            src={pContent.content}
            alt={t('pageSketchAlt')}
            className="w-full h-full object-contain opacity-90 group-hover:opacity-100 transition-opacity"
            style={{ backgroundColor: theme.colors.bookInner }}
          />
          <div className="absolute bottom-12 right-12 text-[9px] font-sans font-bold uppercase tracking-[0.3em] opacity-15" style={{ color: theme.colors.text }}>
            {capDate} · {t('pageSketchCaption')}
          </div>
        </div>
      </div>
    );
  }

  /* ── empty page ── */
  if (pContent.type === 'empty') {
    return (
      <div
        className="w-1/2 h-full min-h-0 p-8 flex items-center justify-center font-serif"
        style={{
          backgroundColor: theme.colors.bookInner,
          ...(side === 'right' ? { paddingLeft: '2.5rem', paddingRight: '2rem', borderLeft: `1px solid ${theme.colors.border}` } : {}),
        }}
      >
        <p className="italic tracking-widest uppercase text-[10px] opacity-20" style={{ color: theme.colors.textMuted }}>{t('pageBlank')}</p>
      </div>
    );
  }

  /* ── text page ── */
  const isRightSecondary = side === 'right' && pContent.secondaryPage !== undefined;
  const hasFallback = !!pContent.leftFallbackSection;
  const sectionTitle = isRightSecondary
    ? (activeSection === 'stt' ? t('pageVoiceTranscript') : t('pageIntelAnalysis'))
    : hasFallback
      ? (pContent.leftFallbackSection === 'stt' ? t('pageVoiceTranscript') : t('pageIntelAnalysis'))
      : t('pageDailyReflection');

  const bigHeaderTitle = hasFallback
    ? (pContent.leftFallbackSection === 'stt' ? t('pageVoiceTranscript') : t('pageIntelAnalysis'))
    : t('pageJournalEntry');

  const pageLabel = isRightSecondary ? pContent.secondaryPage : pContent.journalPage;
  const columnDate = pContent.sourceDate ?? spread.date;

  const showBigHeader = pContent.type === 'text' && !!pContent.showJournalEntryHeader;
  const showSubHeader = !showBigHeader && pContent.type === 'text' && pageLabel !== undefined;

  const layerZIndex = (kind: 'text' | 'sketch' | 'images') => {
    if (!overlay) return 0;
    const order = overlay.layerOrder ?? ['text', 'sketch', 'images'];
    return order.indexOf(kind) + 1;
  };

  return (
    <div
      className="w-1/2 h-full min-h-0 p-8 pr-10 flex flex-col relative overflow-hidden font-serif transition-colors duration-500"
      style={{
        backgroundColor: theme.colors.bookInner,
        color: theme.colors.text,
        ...(side === 'right'
          ? { paddingLeft: '2.5rem', paddingRight: '2rem' }
          : { borderRight: `1px solid ${theme.colors.border}` }),
      }}
    >
      {showBigHeader && (
        <div className="pb-4 mb-6 relative" style={{ borderBottom: `1px solid ${theme.colors.border}` }}>
          <span className="text-xs uppercase tracking-widest font-sans opacity-40" style={{ color: theme.colors.textMuted }}>
            {columnDate} · {pContent.sourceTime ?? spread.time}
          </span>
          <h2 className="text-2xl font-light leading-tight mt-1 font-serif" style={{ color: theme.colors.text }}>
            {bigHeaderTitle}
          </h2>
        </div>
      )}

      {showSubHeader && (
        <div className="pb-2 mb-6 flex justify-between items-end border-dashed relative" style={{ borderBottom: `1px dashed ${theme.colors.border}` }}>
          <h3 className="text-sm font-light font-serif italic tracking-wide" style={{ color: theme.colors.textMuted }}>
            {pageLabel !== undefined && pageLabel > 1
              ? `${sectionTitle} ${t('pageContSuffix')}`
              : sectionTitle}
          </h3>
          <span className="text-[10px] font-sans opacity-30" style={{ color: theme.colors.textMuted }}>{columnDate}</span>
        </div>
      )}

      <div className="flex-1 relative min-h-0 overflow-hidden">
        {/* Text content */}
        <div
          className={`absolute inset-0 leading-relaxed text-base overflow-y-auto pr-2 ${overlay ? 'whitespace-pre-wrap' : 'space-y-3'}`}
          style={{ color: theme.colors.text, zIndex: overlay ? layerZIndex('text') : 'auto' }}
        >
          {overlay
            ? (pContent.content || '')
            : (pContent.content || '').split('\n').map((para, i) => <p key={i}>{para}</p>)
          }
        </div>
      </div>

      {/* Overlay sketch layer – rendered at full page level to match editor canvas coordinates */}
      {overlay?.sketchDataUrl && (
        <img
          src={overlay.sketchDataUrl}
          alt=""
          className="absolute inset-0 w-full h-full pointer-events-none"
          style={{ zIndex: layerZIndex('sketch') }}
        />
      )}

      {/* Overlay image layer – rendered at full page level */}
      {overlay && overlay.images.length > 0 && (
        <div className="absolute inset-0 pointer-events-none" style={{ zIndex: layerZIndex('images') }}>
          {overlay.images.map((img) => (
            <img
              key={img.id}
              src={img.dataUrl}
              alt=""
              className="absolute object-contain"
              style={{
                left: `${img.x * 100}%`,
                top: `${img.y * 100}%`,
                width: `${img.width * 100}%`,
                height: `${img.height * 100}%`,
              }}
            />
          ))}
        </div>
      )}

      <div className="mt-8 flex justify-center shrink-0">
        <div className="w-16 h-1 rounded-full" style={{ backgroundColor: theme.colors.border }} />
      </div>
    </div>
  );
};

/* ──── Bookmark tab on the right edge of the book ──── */

const BookmarkTab: React.FC<{
  label: string;
  tabNumber: string;
  icon: React.ReactNode;
  isActive: boolean;
  onClick: () => void;
  colors: { bg: string; active: string };
}> = ({ label, tabNumber, icon, isActive, onClick, colors }) => {
  return (
    <motion.button
      type="button"
      onClick={onClick}
      className={cn(
        'px-6 py-3 rounded-r-md shadow-sm border-l-4 flex flex-col cursor-pointer transition-all w-36 items-start font-sans',
        isActive ? 'translate-x-2' : 'hover:translate-x-1',
      )}
      style={{
        backgroundColor: isActive ? colors.active : colors.bg,
        color: isActive ? '#fff' : 'inherit',
        borderLeftColor: isActive ? colors.active : `${colors.active}40`,
        opacity: isActive ? 1 : 0.85,
      }}
      whileTap={{ scale: 0.98 }}
    >
      <span className="text-[10px] font-bold uppercase tracking-tighter opacity-60">{tabNumber}</span>
      <span className="text-[11px] font-bold tracking-tight uppercase flex items-center gap-2 whitespace-nowrap">
        {icon}
        {label}
      </span>
    </motion.button>
  );
};

export default JournalBook;
