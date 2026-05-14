import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { createPortal } from 'react-dom';
import { motion, AnimatePresence } from 'motion/react';
import { cn, JournalEntry, JournalSection, PositionedSketch, PageImage, PageOverlay } from '../lib/utils';
import Cover from './Cover';
import Navigation from './Navigation';
import { DrawingCanvas } from './DrawingCanvas';
import { SketchPlacer } from './SketchPlacer';
import { EditorSidebar } from './EditorSidebar';
import { useInlineEditor, type LayerKind } from '../hooks/useInlineEditor';
import { useIsMobile } from '../hooks/useIsMobile';
import { ChevronLeft, ChevronRight, Type, MessageSquare, BrainCircuit, X, GripVertical, Pencil } from 'lucide-react';
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
  bookmarkContent?: PageContent;
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
    // STT or AI mode: same as journal but insert a bookmark page after
    // each entry that has speech/AI content.  Pair everything 2-at-a-time.
    const secondaryField = activeSection === 'stt' ? 'speechToText' : 'aiReport';
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

      const secondaryText = (entry[secondaryField] || '').trim();
      if (secondaryText) {
        items.push({
          content: {
            type: 'text',
            content: pageBreakToNewline(secondaryText),
            secondaryPage: 1,
            sourceEntryId: entry.id,
            sourceDate: displayDate,
            sourceTime: displayTime,
          },
          entryId: entry.id,
          date: displayDate,
          time: displayTime,
        });
      }
    }

    // Pair items into spreads. Bookmark pages (secondaryPage) always go
    // on the RIGHT side of a spread.
    const isBm = (it: PageItem) => !!it.content.secondaryPage;
    let i = 0;
    while (i < items.length) {
      const item = items[i];

      if (item.sketchContents && item.sketchContents.length > 0) {
        spreads.push({
          entryId: item.entryId,
          date: item.date,
          time: item.time,
          left: item.content,
          right: item.sketchContents[0],
          isFirstSpread: true,
        });
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
      } else if (isBm(item)) {
        // Current item is a bookmark — put it on the right, show the
        // entry's journal text on the left so it doesn't disappear.
        let journalLeft: PageContent = { type: 'empty' };
        for (let j = i - 1; j >= 0; j--) {
          if (items[j].entryId === item.entryId && !isBm(items[j])) {
            journalLeft = items[j].content;
            break;
          }
        }
        spreads.push({
          entryId: item.entryId,
          date: item.date,
          time: item.time,
          left: journalLeft,
          right: item.content,
          isFirstSpread: false,
        });
        i += 1;
      } else {
        const nextItem = i + 1 < items.length ? items[i + 1] : null;
        if (nextItem && isBm(nextItem)) {
          // Next item is a bookmark — pair journal (left) + bookmark (right)
          spreads.push({
            entryId: item.entryId,
            date: item.date,
            time: item.time,
            left: item.content,
            right: nextItem.content,
            isFirstSpread: true,
          });
          i += 2;
        } else {
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
  const { isMobile, isLandscape } = useIsMobile();
  const [forceSinglePage, setForceSinglePage] = useState(false);
  const singlePage = (isMobile && !isLandscape) || forceSinglePage;
  const [entries, setEntries] = useState<JournalEntry[]>([]);
  const [sketches, setSketches] = useState<PositionedSketch[]>([]);
  const [overlays, setOverlays] = useState<Record<string, PageOverlay>>({});
  const [appTitle, setAppTitle] = useState('Daily Logger');
  const [loadError, setLoadError] = useState<string | null>(null);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [currentPage, setCurrentPage] = useState(0);
  const [mobileSide, setMobileSide] = useState<'left' | 'right'>('left');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>(() => readStored(READER_SORT_KEY, ['asc', 'desc'], 'asc'));
  const [activeSection, setActiveSection] = useState<JournalSection>('journal');

  const [showSketchPlacer, setShowSketchPlacer] = useState(false);
  const [editingSketchId, setEditingSketchId] = useState<string | null>(null);

  const [inlineEditEntry, setInlineEditEntry] = useState<{ entryId: string; defaultLayer: LayerKind; side: 'left' | 'right' } | null>(null);

  const bookSpreadRef = useRef<HTMLDivElement>(null);
  const leftPageRef = useRef<HTMLDivElement>(null);
  const rightPageRef = useRef<HTMLDivElement>(null);

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
    fetch('/api/reader-settings')
      .then((r) => r.json())
      .then((data) => {
        if (data.sortOrder === 'asc' || data.sortOrder === 'desc') {
          setSortOrder(data.sortOrder);
          persistStored(READER_SORT_KEY, data.sortOrder);
        }
        if (typeof data.singlePageMode === 'boolean') {
          setForceSinglePage(data.singlePageMode);
        }
      })
      .catch(() => {});
  }, []);

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
    () => buildUnifiedSpreads(sortedEntries, sketches, singlePage ? 'journal' : activeSection),
    [sortedEntries, sketches, activeSection, singlePage],
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
    if (singlePage) {
      if (activeSection !== 'journal') setActiveSection('journal');
      if (mobileSide === 'left') {
        const spread = currentPage > 0 ? spreads[currentPage - 1] : undefined;
        if (spread && spread.right.type !== 'empty') {
          setMobileSide('right');
          return;
        }
      }
      if (currentPage >= spreadCount) return;
      setCurrentPage((p) => p + 1);
      setMobileSide('left');
    } else {
      if (currentPage >= spreadCount) return;
      setCurrentPage((p) => p + 1);
    }
  }, [currentPage, spreadCount, singlePage, mobileSide, spreads, activeSection]);

  const handlePrev = useCallback(() => {
    if (singlePage) {
      if (activeSection !== 'journal') setActiveSection('journal');
      if (mobileSide === 'right') {
        setMobileSide('left');
        return;
      }
      if (currentPage <= 0) return;
      const prevPage = currentPage - 1;
      if (prevPage > 0) {
        const prevSpread = spreads[prevPage - 1];
        setCurrentPage(prevPage);
        setMobileSide(prevSpread && prevSpread.right.type !== 'empty' ? 'right' : 'left');
      } else {
        setCurrentPage(0);
        setMobileSide('left');
      }
    } else {
      if (currentPage <= 0) return;
      setCurrentPage((p) => Math.max(0, p - 1));
    }
  }, [currentPage, singlePage, mobileSide, spreads]);

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
      if (inlineEditEntry) return;
      const key = e.key.toLowerCase();
      if (key === 'd' || key === 'arrowright') handleNext();
      if (key === 'a' || key === 'arrowleft') handlePrev();
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleNext, handlePrev, inlineEditEntry]);

  /* ── swipe nav (mobile) ── */

  const touchStartX = useRef(0);
  const touchStartY = useRef(0);

  const onTouchStart = useCallback((e: React.TouchEvent) => {
    touchStartX.current = e.touches[0].clientX;
    touchStartY.current = e.touches[0].clientY;
  }, []);

  const onTouchEnd = useCallback((e: React.TouchEvent) => {
    if (inlineEditEntry) return;
    const dx = e.changedTouches[0].clientX - touchStartX.current;
    const dy = e.changedTouches[0].clientY - touchStartY.current;
    if (Math.abs(dx) < 50 || Math.abs(dy) > Math.abs(dx)) return;
    if (dx < 0) handleNext();
    else handlePrev();
  }, [handleNext, handlePrev, inlineEditEntry]);

  /* ── mobile page count (individual pages vs spreads) ── */

  const mobilePageTotal = useMemo(() => {
    if (!singlePage) return spreadCount;
    let count = 0;
    for (const s of spreads) {
      count += 1;
      if (s.right.type !== 'empty') count += 1;
    }
    return count;
  }, [singlePage, spreads, spreadCount]);

  const mobilePageCurrent = useMemo(() => {
    if (!singlePage || currentPage === 0) return currentPage;
    let count = 0;
    for (let i = 0; i < currentPage - 1 && i < spreads.length; i++) {
      count += 1;
      if (spreads[i].right.type !== 'empty') count += 1;
    }
    count += 1;
    if (mobileSide === 'right') count += 1;
    return count;
  }, [singlePage, currentPage, mobileSide, spreads]);

  /* ── actions ── */

  const handleAction = (action: JournalAction) => {
    if (action === 'sketch') {
      setShowSketchPlacer(true);
    } else {
      if (currentPage === 0) return;
      const spread = spreads[currentPage - 1];
      if (!spread) return;
      const leftEid = spread.left.sourceEntryId ?? spread.entryId;
      if (leftEid) {
        setInlineEditEntry({ entryId: leftEid, defaultLayer: 'text', side: 'left' });
      }
    }
  };

  /* ── inline editor ── */

  const handleOpenInlineEditor = (entryId: string, defaultLayer: 'text' | 'sketch') => {
    setInlineEditEntry({ entryId, defaultLayer, side: 'left' });
  };

  const doSaveEntry = async (
    entryId: string,
    text: string,
    sketchDataUrl: string,
    images: PageImage[],
    layerOrder: ('text' | 'sketch' | 'images')[],
    date?: string,
    time?: string,
  ): Promise<boolean> => {
    setSaveError(null);
    try {
      const entryBody: Record<string, string> = { id: entryId, journal: text };
      if (date !== undefined) entryBody.date = date;
      if (time !== undefined) entryBody.time = time;
      const res = await fetch('/api/entry', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(entryBody),
      });
      const data = await res.json();
      if (!data.ok) { setSaveError(data.error || t('errSaveFailed')); return false; }

      const overlayBody = { entryId, sketchDataUrl, images, layerOrder };
      const res2 = await fetch('/api/page-overlay', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(overlayBody),
      });
      const data2 = await res2.json();
      if (!data2.ok) { setSaveError(data2.error || t('errSaveFailed')); return false; }

      await reload();
      return true;
    } catch { setSaveError(t('errNetworkSave')); return false; }
  };

  const handleSaveInline = async (
    text: string,
    sketchDataUrl: string,
    images: PageImage[],
    layerOrder: ('text' | 'sketch' | 'images')[],
    date?: string,
    time?: string,
  ) => {
    if (!inlineEditEntry) return;
    const ok = await doSaveEntry(inlineEditEntry.entryId, text, sketchDataUrl, images, layerOrder, date, time);
    if (ok) setInlineEditEntry(null);
  };

  const handleSwitchSide = async (
    currentPayload: { entryId: string; text: string; sketchDataUrl: string; images: PageImage[]; layerOrder: LayerKind[]; activeLayer: LayerKind; date: string; time: string },
    targetEntryId: string,
    targetSide: 'left' | 'right',
  ) => {
    await doSaveEntry(currentPayload.entryId, currentPayload.text, currentPayload.sketchDataUrl, currentPayload.images, currentPayload.layerOrder, currentPayload.date, currentPayload.time);
    setInlineEditEntry({ entryId: targetEntryId, defaultLayer: currentPayload.activeLayer, side: targetSide });
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

  /* ── bookmark page editing (speech/AI) ── */

  const [bookmarkEdit, setBookmarkEdit] = useState<{ entryId: string; field: 'speechToText' | 'aiReport'; text: string } | null>(null);

  const handleEditBookmarkPage = (entryId: string, field: 'speechToText' | 'aiReport', currentText: string) => {
    setBookmarkEdit({ entryId, field, text: currentText });
  };

  const handleSaveBookmarkEdit = async () => {
    if (!bookmarkEdit) return;
    setSaveError(null);
    try {
      const body: Record<string, string> = { id: bookmarkEdit.entryId };
      body[bookmarkEdit.field] = bookmarkEdit.text;
      const res = await fetch('/api/entry', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      const data = await res.json();
      if (!data.ok) { setSaveError(data.error || t('errSaveFailed')); return; }
      await reload();
      setBookmarkEdit(null);
    } catch { setSaveError(t('errNetworkSave')); }
  };

  const handleBookmarkClick = (section: JournalSection) => {
    if (section !== 'journal' && inlineEditEntry) {
      setInlineEditEntry(null);
    }
    if (singlePage) {
      setActiveSection(section);
      return;
    }
    // Remember the entry the user is currently viewing so we can stay on
    // it after the spreads rebuild with a different page count.
    const viewingEntryId = currentSpread?.entryId;
    setActiveSection(section);
    if (viewingEntryId) {
      const newSpreads = buildUnifiedSpreads(sortedEntries, sketches, section);
      let ix = firstSpreadIndexForEntry(newSpreads, viewingEntryId);
      if (ix < 0) ix = newSpreads.findIndex((s) => spreadTouchesEntryId(s, viewingEntryId));
      if (ix >= 0) setCurrentPage(ix + 1);
    }
  };

  const currentSpread = currentPage > 0 ? spreads[currentPage - 1] : undefined;

  /* In single-page mode, figure out the entry shown on the current page
     and whether it has STT/AI content.  When a bookmark tab is active,
     we swap the page content instead of rebuilding spreads. */
  const singlePageEntry = useMemo(() => {
    if (!singlePage || !currentSpread) return undefined;
    const page = mobileSide === 'left' ? currentSpread.left : currentSpread.right;
    const eid = page.sourceEntryId ?? currentSpread.entryId;
    return entries.find((e) => e.id === eid);
  }, [singlePage, currentSpread, mobileSide, entries]);

  const singlePageHasStt = !!(singlePageEntry?.speechToText?.trim());
  const singlePageHasAi = !!(singlePageEntry?.aiReport?.trim());
  const singlePageHasBookmark = singlePageHasStt || singlePageHasAi;

  const singlePageBookmarkContent = useMemo<PageContent | undefined>(() => {
    if (!singlePage || !singlePageEntry || activeSection === 'journal') return undefined;
    const field = activeSection === 'stt' ? 'speechToText' : 'aiReport';
    const text = (singlePageEntry[field] || '').trim();
    if (!text) return undefined;
    return {
      type: 'text',
      content: text,
      secondaryPage: 1,
      sourceEntryId: singlePageEntry.id,
      sourceDate: singlePageEntry.date,
      sourceTime: singlePageEntry.time,
    };
  }, [singlePage, singlePageEntry, activeSection]);

  return (
    <div className="flex h-dvh max-h-dvh min-h-0 flex-col overflow-hidden select-none font-sans transition-colors duration-500" style={{ backgroundColor: bgTheme.colors.bg }}>
      {!singlePage && (
        <div className="absolute top-6 left-8 opacity-80 z-10" style={{ color: bgTheme.cover.accentText }}>
          <h1 className="text-2xl tracking-widest uppercase font-light font-serif">
            {appTitle}{' '}
            <span className="text-xs opacity-50 block tracking-normal font-sans">{t('readerSubtitle')}</span>
          </h1>
        </div>
      )}

      {(loadError || saveError) && (
        <div className="fixed top-6 left-1/2 -translate-x-1/2 z-[200] max-w-xl rounded-lg border border-amber-500/40 bg-black/80 px-6 py-3 text-sm text-amber-100 font-sans shadow-2xl backdrop-blur-sm">
          {loadError && <p>{loadError}</p>}
          {saveError && <p>{saveError}</p>}
        </div>
      )}

      <Navigation
        currentPage={singlePage ? mobilePageCurrent : currentPage}
        totalPages={singlePage ? mobilePageTotal : spreadCount}
        onPageJump={handlePageJump}
        onPrev={handlePrev}
        onNext={handleNext}
        onAction={handleAction}
        onToggleSort={() =>
          setSortOrder((prev) => {
            const next = prev === 'asc' ? 'desc' : 'asc';
            persistStored(READER_SORT_KEY, next);
            fetch('/api/reader-settings', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ sortOrder: next }),
            }).catch(() => {});
            return next;
          })
        }
        sortOrder={sortOrder}
        onDateSelect={handleDateSelect}
        isEditTextOpen={false}
        onSaveText={() => {}}
        isMobile={singlePage}
        singlePageMode={forceSinglePage}
        onToggleSinglePage={() => {
          const next = !forceSinglePage;
          setForceSinglePage(next);
          setMobileSide('left');
          fetch('/api/reader-settings', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ singlePageMode: next }),
          }).catch(() => {});
        }}
      />

      <div
        className="flex-1 flex min-h-0 items-center justify-center p-2 md:p-5 perspective-[2000px]"
        onTouchStart={onTouchStart}
        onTouchEnd={onTouchEnd}
        style={{ touchAction: 'manipulation' }}
      >
        <div
          className={cn(
            'relative mx-auto flex h-auto max-h-full min-h-0 max-w-full shrink rounded-xl',
            singlePage
              ? 'aspect-[0.7/1] w-[min(100%,24rem,92vw)] shadow-[0_20px_60px_-10px_rgba(0,0,0,0.5)]'
              : 'aspect-[1.4/1] w-[min(100%,72rem,92vw)] shadow-[0_50px_100px_-20px_rgba(0,0,0,0.5)]',
          )}
          style={{ transform: inlineEditEntry && !singlePage ? 'translateX(7rem)' : 'none', transition: 'transform 0.3s ease' }}
        >
          <div className="absolute inset-0 bg-black/40 -z-10 rounded-xl translate-x-2 translate-y-2 blur-2xl" />

          {!singlePage && (
            <div
              className={cn(
                'absolute left-1/2 -ml-0.5 top-0 bottom-0 w-px z-20 shadow-[0_0_10px_rgba(0,0,0,0.1)]',
                currentPage === 0 && 'hidden',
              )}
              style={{ backgroundColor: bgTheme.colors.spine }}
            />
          )}

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

                {/* Click zones for page flip – hidden during inline editing */}
                {!inlineEditEntry && (
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
                )}

                {singlePage ? (
                  singlePageBookmarkContent ? (
                    <Page
                      ref={leftPageRef}
                      spread={currentSpread ? {
                        ...currentSpread,
                        left: singlePageBookmarkContent,
                        right: { type: 'empty' },
                      } : undefined}
                      side="left"
                      activeSection={activeSection}
                      theme={bgTheme}
                      fullWidth
                      onEditBookmark={handleEditBookmarkPage}
                    />
                  ) : (
                    <Page
                      ref={mobileSide === 'left' ? leftPageRef : rightPageRef}
                      spread={currentSpread}
                      side={mobileSide}
                      activeSection={activeSection}
                      theme={bgTheme}
                      overlay={(() => {
                        const eid = mobileSide === 'left' ? currentSpread?.left.sourceEntryId : currentSpread?.right.sourceEntryId;
                        return eid ? overlays[eid] : undefined;
                      })()}
                      fullWidth
                      onEditBookmark={handleEditBookmarkPage}
                    />
                  )
                ) : (
                  <>
                    <Page
                      ref={leftPageRef}
                      spread={currentSpread}
                      side="left"
                      activeSection={activeSection}
                      theme={bgTheme}
                      overlay={currentSpread?.left.sourceEntryId ? overlays[currentSpread.left.sourceEntryId] : undefined}
                      onEditBookmark={handleEditBookmarkPage}
                    />
                    <Page
                      ref={rightPageRef}
                      spread={currentSpread}
                      side="right"
                      activeSection={activeSection}
                      theme={bgTheme}
                      overlay={currentSpread?.right.sourceEntryId && currentSpread.right.sourceEntryId !== currentSpread.left.sourceEntryId ? overlays[currentSpread.right.sourceEntryId] : undefined}
                      onEditBookmark={handleEditBookmarkPage}
                    />
                  </>
                )}

                {/* Inline editor layers inside the book spread — absolute positioning for zoom-proof alignment.
                   Only the in-book layers go here; sidebar + file input are rendered outside (see below)
                   because motion.div's transform creates a containing block that traps fixed elements. */}
                {inlineEditEntry && (() => {
                  const editEntry = entries.find((e) => e.id === inlineEditEntry.entryId);
                  const editOv = overlays[inlineEditEntry.entryId];
                  const editOtherSide = inlineEditEntry.side === 'left' ? 'right' : 'left';
                  const editOtherPage = currentSpread ? (editOtherSide === 'left' ? currentSpread.left : currentSpread.right) : undefined;
                  const editOtherEntryId = editOtherPage?.sourceEntryId ?? currentSpread?.entryId;
                  return (
                    <InlineEditorOverlay
                      key={`${inlineEditEntry.entryId}-${inlineEditEntry.side}`}
                      entry={editEntry}
                      overlay={editOv}
                      defaultLayer={inlineEditEntry.defaultLayer}
                      side={inlineEditEntry.side}
                      pageRef={inlineEditEntry.side === 'left' ? leftPageRef : rightPageRef}
                      otherEntryId={singlePage ? undefined : editOtherEntryId}
                      otherSide={editOtherSide}
                      onSave={handleSaveInline}
                      onSwitchSide={handleSwitchSide}
                      onClose={() => setInlineEditEntry(null)}
                      isMobile={singlePage}
                    />
                  );
                })()}
              </motion.div>
            )}
          </AnimatePresence>

          {/* Bookmark tabs */}
          {currentPage > 0 && (singlePage ? singlePageHasBookmark : true) && (
            <div className={cn(
              'absolute z-30',
              singlePage
                ? 'left-0 right-0 top-full mt-1 flex flex-row justify-center space-x-1'
                : 'left-full ml-1 top-20 flex flex-col space-y-1',
            )}>
              {/* Journal tab — in single page only show when viewing a bookmark */}
              {(!singlePage || activeSection !== 'journal') && (
                <BookmarkTab
                  label={t('tabJournal')}
                  tabNumber="01"
                  icon={<Type size={singlePage ? 14 : 16} />}
                  isActive={activeSection === 'journal'}
                  onClick={() => handleBookmarkClick('journal')}
                  colors={bgTheme.colors.tabs.journal}
                  compact={singlePage}
                />
              )}
              {(!singlePage || singlePageHasStt) && (
                <BookmarkTab
                  label={t('tabSpeech')}
                  tabNumber="02"
                  icon={<MessageSquare size={singlePage ? 14 : 16} />}
                  isActive={activeSection === 'stt'}
                  onClick={() => handleBookmarkClick('stt')}
                  colors={bgTheme.colors.tabs.stt}
                  compact={singlePage}
                />
              )}
              {(!singlePage || singlePageHasAi) && (
                <BookmarkTab
                  label={t('tabAi')}
                  tabNumber="03"
                  icon={<BrainCircuit size={singlePage ? 14 : 16} />}
                  isActive={activeSection === 'ai'}
                  onClick={() => handleBookmarkClick('ai')}
                  colors={bgTheme.colors.tabs.ai}
                  compact={singlePage}
                />
              )}
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
          onOpenCompositeEditor={handleOpenInlineEditor}
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

      {/* Bookmark page edit modal (speech/AI) */}
      {bookmarkEdit && (
        <div className="fixed inset-0 z-[200] bg-black/60 flex items-center justify-center p-4 backdrop-blur-sm font-sans">
          <div
            className="rounded-2xl shadow-2xl w-full max-w-2xl flex flex-col overflow-hidden"
            style={{ backgroundColor: bgTheme.colors.bookInner, maxHeight: '80vh' }}
          >
            <div className="flex items-center justify-between px-4 md:px-6 py-3 shrink-0" style={{ borderBottom: `1px solid ${bgTheme.colors.border}` }}>
              <h3 className="font-semibold uppercase tracking-widest text-xs" style={{ color: bgTheme.colors.text }}>
                {bookmarkEdit.field === 'speechToText' ? t('tabSpeech') : t('tabAi')}
              </h3>
              <div className="flex items-center gap-2">
                <button
                  onClick={handleSaveBookmarkEdit}
                  className="px-4 py-1.5 rounded-lg text-white text-xs font-semibold uppercase tracking-widest min-h-[36px]"
                  style={{ backgroundColor: bgTheme.cover.isDark ? '#4f46e5' : '#334155' }}
                >
                  {t('compositeEditorSave')}
                </button>
                <button
                  onClick={() => setBookmarkEdit(null)}
                  className="p-1.5 rounded-full hover:bg-black/5 transition-colors"
                  style={{ color: bgTheme.colors.textMuted }}
                >
                  <X size={18} />
                </button>
              </div>
            </div>
            <textarea
              className="flex-1 min-h-[200px] p-4 md:p-6 resize-none focus:outline-none font-serif text-base leading-relaxed"
              style={{ backgroundColor: bgTheme.colors.bookInner, color: bgTheme.colors.text }}
              value={bookmarkEdit.text}
              onChange={(e) => setBookmarkEdit((prev) => prev ? { ...prev, text: e.target.value } : prev)}
              autoFocus
            />
          </div>
        </div>
      )}

      <footer className="shrink-0 flex flex-col items-center gap-1 md:gap-2 px-4 pb-2 md:pb-4 pt-1 md:pt-2 text-center" style={{ color: bgTheme.cover.accentText }}>
        <div className={cn('flex items-center justify-center opacity-40', singlePage ? 'space-x-6' : 'space-x-12')}>
          <button
            type="button"
            onClick={handlePrev}
            className="flex items-center space-x-1 md:space-x-2 group hover:opacity-100 transition-opacity min-h-[44px] min-w-[44px] justify-center"
          >
            <ChevronLeft className="w-5 h-5 group-hover:-translate-x-1 transition-transform" />
            {!singlePage && <span className="text-sm uppercase tracking-widest font-sans">{t('footerPrev')}</span>}
          </button>
          <div className="flex items-center space-x-3">
            <div className={cn('w-1 h-1 rounded-full', currentPage < 2 ? 'bg-current/50' : 'bg-current/20')} />
            <div className={cn('w-1.5 h-1.5 rounded-full', currentPage >= 2 && currentPage < spreadCount ? 'bg-current/50' : 'bg-current/20')} />
            <div className={cn('w-1 h-1 rounded-full', currentPage >= spreadCount ? 'bg-current/50' : 'bg-current/20')} />
          </div>
          <button
            type="button"
            onClick={handleNext}
            className="flex items-center space-x-1 md:space-x-2 group hover:opacity-100 transition-opacity min-h-[44px] min-w-[44px] justify-center"
          >
            {!singlePage && <span className="text-sm uppercase tracking-widest font-sans">{t('footerNext')}</span>}
            <ChevronRight className="w-5 h-5 group-hover:translate-x-1 transition-transform" />
          </button>
        </div>
        {!singlePage && (
          <div className="text-[10px] uppercase tracking-[0.3em] opacity-20 font-sans leading-relaxed">
            {t('footerFlipHint')}
          </div>
        )}
      </footer>
    </div>
  );
};

/* ──── Inline editor overlay — sidebar + canvas/textarea/images on the active page ──── */

const InlineEditorOverlay: React.FC<{
  entry?: JournalEntry;
  overlay?: PageOverlay;
  defaultLayer: LayerKind;
  side: 'left' | 'right';
  pageRef: React.RefObject<HTMLDivElement | null>;
  otherEntryId?: string;
  otherSide: 'left' | 'right';
  onSave: (text: string, sketchDataUrl: string, images: PageImage[], layerOrder: LayerKind[], date?: string, time?: string) => void;
  onSwitchSide: (
    currentPayload: { entryId: string; text: string; sketchDataUrl: string; images: PageImage[]; layerOrder: LayerKind[]; activeLayer: LayerKind; date: string; time: string },
    targetEntryId: string,
    targetSide: 'left' | 'right',
  ) => void;
  onClose: () => void;
  isMobile?: boolean;
}> = ({ entry, overlay: ov, defaultLayer, side, pageRef, otherEntryId, otherSide, onSave, onSwitchSide, onClose, isMobile }) => {
  const { t } = useReaderT();
  const { bgTheme } = useTheme();
  const editor = useInlineEditor({
    initialText: entry?.journal ?? '',
    initialSketchDataUrl: ov?.sketchDataUrl,
    initialImages: ov?.images ?? [],
    initialLayerOrder: ov?.layerOrder ?? ['text', 'sketch', 'images'],
    defaultLayer,
    pageAreaRef: pageRef,
  });

  const [editDate, setEditDate] = React.useState(entry?.date ?? '');
  const [editTime, setEditTime] = React.useState(entry?.time ?? '');

  const handleSave = () => {
    const payload = editor.getSavePayload();
    onSave(payload.text, payload.sketchDataUrl, payload.images, payload.layerOrder, editDate, editTime);
  };

  const handleClickOtherPage = () => {
    if (!otherEntryId || !entry) return;
    const payload = editor.getSavePayload();
    onSwitchSide(
      { entryId: entry.id, text: payload.text, sketchDataUrl: payload.sketchDataUrl, images: payload.images, layerOrder: payload.layerOrder, activeLayer: editor.activeLayer, date: editDate, time: editTime },
      otherEntryId,
      otherSide,
    );
  };

  const padClass = isMobile ? 'p-4' : (side === 'right' ? 'p-8 pl-10' : 'p-8 pr-10');
  const overlayWidth = isMobile ? 'w-full' : 'w-1/2';
  const pagePosition = isMobile ? 'left-0' : (side === 'left' ? 'left-0' : 'left-1/2');

  return (
    <>
      {/* Sidebar + file input portalled to document.body so they escape the
          motion.div's CSS transform (which creates a containing block and
          traps position:fixed elements inside the book spread). */}
      {createPortal(
        <>
          <EditorSidebar
            activeLayer={editor.activeLayer}
            onLayerChange={editor.setActiveLayer}
            layerOrder={editor.layerOrder}
            onMoveLayer={editor.moveLayer}
            onReorderLayers={editor.reorderLayers}
            color={editor.color}
            lineWidth={editor.lineWidth}
            isErasing={editor.isErasing}
            eraserSize={editor.eraserSize}
            onColorChange={editor.setColor}
            onLineWidthChange={editor.setLineWidth}
            onSetErasing={editor.setIsErasing}
            onEraserSizeChange={editor.setEraserSize}
            onClearCanvas={editor.clearCanvas}
            onUploadImage={editor.handleUpload}
            onSave={handleSave}
            onClose={onClose}
            entryDate={editDate}
            entryTime={editTime}
            isMobile={isMobile}
          />
          <input ref={editor.fileInputRef} type="file" accept="image/*" className="hidden" onChange={editor.handleFileChange} />
        </>,
        document.body,
      )}

      {/* Click-to-switch overlay on the opposite page */}
      {otherEntryId && (
        <div
          className={`absolute top-0 bottom-0 ${overlayWidth} z-[60] cursor-pointer hover:bg-blue-400/10 transition-colors ${otherSide === 'left' ? 'left-0' : 'left-1/2'}`}
          onClick={handleClickOtherPage}
          title="Click to edit this page"
        />
      )}

      {/* Sketch canvas — absolute within book spread, covers the active page */}
      <canvas
        ref={editor.canvasRef}
        className={`absolute top-0 bottom-0 ${overlayWidth} touch-none ${pagePosition}`}
        style={{
          zIndex: 50 + editor.zIndex('sketch'),
          pointerEvents: editor.activeLayer === 'sketch' ? 'auto' : 'none',
          cursor: editor.activeLayer === 'sketch' ? editor.eraserCursor : 'default',
        }}
        onMouseDown={editor.startDrawing}
        onMouseMove={editor.draw}
        onMouseUp={editor.stopDrawing}
        onMouseOut={editor.stopDrawing}
        onTouchStart={editor.startDrawing}
        onTouchMove={editor.draw}
        onTouchEnd={editor.stopDrawing}
      />

      {/* Text + image editing layers — absolute within book spread, covers the active page */}
      <div className={`absolute top-0 bottom-0 ${overlayWidth} ${pagePosition}`} style={{ zIndex: 50 }}>
        {/* Text layer — opaque background to hide rendered page text underneath */}
        <div
          className={`absolute inset-0 ${padClass} flex flex-col font-serif`}
          style={{ zIndex: editor.zIndex('text'), pointerEvents: editor.activeLayer === 'text' ? 'auto' : 'none', backgroundColor: bgTheme.colors.bookInner }}
        >
          <div className="pb-4 mb-6" style={{ borderBottom: `1px solid ${bgTheme.colors.border}` }}>
            <div className="flex items-center gap-2">
              <input
                value={editDate}
                onChange={(e) => setEditDate(e.target.value)}
                className="text-xs uppercase tracking-widest font-sans opacity-60 bg-transparent border-none focus:outline-none focus:opacity-100 transition-opacity w-24"
                style={{ color: bgTheme.colors.textMuted }}
                placeholder="MM/DD/YYYY"
              />
              <span className="text-xs opacity-40" style={{ color: bgTheme.colors.textMuted }}>·</span>
              <input
                value={editTime}
                onChange={(e) => setEditTime(e.target.value)}
                className="text-xs uppercase tracking-widest font-sans opacity-60 bg-transparent border-none focus:outline-none focus:opacity-100 transition-opacity w-20"
                style={{ color: bgTheme.colors.textMuted }}
                placeholder="HH:MM"
              />
            </div>
            <h2 className="text-2xl font-light leading-tight mt-1 font-serif" style={{ color: bgTheme.colors.text }}>
              {t('pageJournalEntry')}
            </h2>
          </div>
          <div className="flex-1 relative min-h-0">
            <textarea
              value={editor.text}
              onChange={(e) => editor.setText(e.target.value)}
              className="absolute inset-0 w-full h-full bg-transparent resize-none font-serif text-base leading-relaxed focus:outline-none overflow-auto pr-2"
              style={{ color: bgTheme.colors.text }}
              placeholder="Type here..."
            />
          </div>
          <div className="mt-8 shrink-0 h-1 pointer-events-none" />
        </div>

        {/* Image layer */}
        <div
          className="absolute inset-0"
          style={{ zIndex: editor.zIndex('images'), pointerEvents: editor.activeLayer === 'images' ? 'auto' : 'none' }}
          onMouseDown={(e) => { if (e.target === e.currentTarget) editor.setSelectedImgId(null); }}
        >
          {editor.images.map((img) => {
            const isSelected = editor.selectedImgId === img.id;
            return (
              <div
                key={img.id}
                className="absolute"
                style={{
                  left: `${img.x * 100}%`,
                  top: `${img.y * 100}%`,
                  width: `${img.width * 100}%`,
                  height: `${img.height * 100}%`,
                  outline: isSelected ? '2px solid #3b82f6' : 'none',
                  cursor: editor.activeLayer === 'images' ? 'move' : 'default',
                }}
                onMouseDown={(e) => editor.handleImageMouseDown(e, img)}
              >
                <img src={img.dataUrl} alt="" className="w-full h-full object-contain select-none pointer-events-none" draggable={false} />
                {isSelected && editor.activeLayer === 'images' && (
                  <>
                    <button
                      className="absolute -top-3 -right-3 w-6 h-6 bg-red-500 text-white rounded-full flex items-center justify-center shadow hover:bg-red-600 transition-colors"
                      onClick={(e) => { e.stopPropagation(); editor.deleteImage(img.id); }}
                    >
                      <X size={14} />
                    </button>
                    <div
                      className="absolute -bottom-2 -right-2 w-5 h-5 bg-blue-500 rounded-sm cursor-se-resize flex items-center justify-center shadow"
                      onMouseDown={(e) => { e.stopPropagation(); editor.handleResizeMouseDown(e, img); }}
                    >
                      <GripVertical size={10} className="text-white rotate-45" />
                    </div>
                  </>
                )}
              </div>
            );
          })}
        </div>

        {/* Subtle editing border */}
        <div className="absolute inset-0 border-2 border-blue-400/30 rounded-lg pointer-events-none" style={{ zIndex: 10 }} />
      </div>
    </>
  );
};

/* ──── Page sub-component ──── */

const Page = React.forwardRef<HTMLDivElement, {
  spread?: Spread;
  side: 'left' | 'right';
  activeSection: JournalSection;
  theme: JournalTheme;
  overlay?: PageOverlay;
  fullWidth?: boolean;
  onEditBookmark?: (entryId: string, field: 'speechToText' | 'aiReport', text: string) => void;
}>(({ spread, side, activeSection, theme, overlay, fullWidth, onEditBookmark }, ref) => {
  const { t } = useReaderT();
  const widthCls = fullWidth ? 'w-full' : 'w-1/2';
  if (!spread) {
    return (
      <div
        className={`${widthCls} h-full min-h-0 p-4 md:p-8 md:pr-10 flex items-center justify-center font-serif`}
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
      <div className={`${widthCls} h-full min-h-0 relative group overflow-hidden flex flex-col`} style={{ backgroundColor: theme.colors.bookInner }}>
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
        className={`${widthCls} h-full min-h-0 p-4 md:p-8 flex items-center justify-center font-serif`}
        style={{
          backgroundColor: theme.colors.bookInner,
          ...(side === 'right' && !fullWidth ? { paddingLeft: '2.5rem', paddingRight: '2rem', borderLeft: `1px solid ${theme.colors.border}` } : {}),
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
      ref={ref}
      className={`${widthCls} h-full min-h-0 p-4 md:p-8 md:pr-10 flex flex-col relative overflow-hidden font-serif transition-colors duration-500`}
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
          <div className="flex items-center gap-2">
            {isRightSecondary && onEditBookmark && pContent.sourceEntryId && (
              <button
                type="button"
                onClick={() => onEditBookmark(
                  pContent.sourceEntryId!,
                  activeSection === 'stt' ? 'speechToText' : 'aiReport',
                  pContent.content || '',
                )}
                className="p-1 rounded-full hover:bg-black/10 transition-colors z-50 min-h-[36px] min-w-[36px] flex items-center justify-center"
                style={{ color: theme.colors.textMuted }}
                title="Edit"
              >
                <Pencil size={14} />
              </button>
            )}
            <span className="text-[10px] font-sans opacity-30" style={{ color: theme.colors.textMuted }}>{columnDate}</span>
          </div>
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

      {/* Overlay sketch layer */}
      {overlay?.sketchDataUrl && (
        <img
          src={overlay.sketchDataUrl}
          alt=""
          className="absolute inset-0 w-full h-full pointer-events-none"
          style={{ zIndex: layerZIndex('sketch') }}
        />
      )}

      {/* Overlay image layer */}
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
});

/* ──── Bookmark tab on the right edge of the book ──── */

const BookmarkTab: React.FC<{
  label: string;
  tabNumber: string;
  icon: React.ReactNode;
  isActive: boolean;
  onClick: () => void;
  colors: { bg: string; active: string };
  compact?: boolean;
}> = ({ label, tabNumber, icon, isActive, onClick, colors, compact }) => {
  if (compact) {
    return (
      <motion.button
        type="button"
        onClick={onClick}
        className={cn(
          'px-3 py-2 rounded-b-md shadow-sm border-t-2 flex items-center gap-1.5 cursor-pointer transition-all font-sans min-h-[40px]',
          isActive && 'translate-y-0.5',
        )}
        style={{
          backgroundColor: isActive ? colors.active : colors.bg,
          color: isActive ? '#fff' : 'inherit',
          borderTopColor: isActive ? colors.active : `${colors.active}40`,
          opacity: isActive ? 1 : 0.85,
        }}
        whileTap={{ scale: 0.97 }}
      >
        {icon}
        <span className="text-[10px] font-bold tracking-tight uppercase whitespace-nowrap">{label}</span>
      </motion.button>
    );
  }

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
