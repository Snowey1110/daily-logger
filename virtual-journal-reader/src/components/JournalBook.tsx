import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { cn, JournalEntry, JournalSection, PositionedSketch, RightPageSetting } from '../lib/utils';
import Cover from './Cover';
import Navigation from './Navigation';
import { DrawingCanvas } from './DrawingCanvas';
import { SketchPlacer } from './SketchPlacer';
import { ChevronLeft, ChevronRight, Type, MessageSquare, BrainCircuit } from 'lucide-react';
import { useReaderT } from '../readerI18n';

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
  rightAltContent?: PageContent;
  rightHasTabs?: boolean;
  rightTabDefault?: 'sketch' | 'secondary';
}

/* ──── localStorage helpers ──── */

const READER_SORT_KEY = 'virtualJournalReader.sortOrder';
const RIGHT_PAGE_KEY = 'virtualJournalReader.rightPageSetting';

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

  const left: PageContent = leftText
    ? {
        type: 'text',
        content: leftText,
        journalPage: 1,
        sourceEntryId: entry.id,
        sourceDate: displayDate,
        sourceTime: displayTime,
        showJournalEntryHeader: true,
        leftFallbackSection: leftFallback !== 'journal' ? leftFallback : undefined,
      }
    : { type: 'empty' };

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
        rightHasTabs: i === 0 ? (spread.rightHasTabs ?? false) : false,
        rightAltContent: i === 0 ? spread.rightAltContent : undefined,
        rightTabDefault: spread.rightTabDefault,
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
  sketchContent?: PageContent;
}

function expandEntryToPages(
  entry: JournalEntry,
  displayDate: string,
  displayTime: string,
  left: PageContent,
  sketchPC?: PageContent,
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
    sketchContent: j === textPages.length - 1 ? sketchPC : undefined,
  }));
}

function buildUnifiedSpreads(
  sortedEntries: JournalEntry[],
  sketches: PositionedSketch[],
  activeSection: JournalSection,
): Spread[] {
  const sketchByEntry = new Map<string, PositionedSketch>();
  for (const sk of sketches) {
    if (!sketchByEntry.has(sk.afterEntryId)) {
      sketchByEntry.set(sk.afterEntryId, sk);
    }
  }

  const spreads: Spread[] = [];

  if (activeSection === 'journal') {
    // Expand all entries into page items (splitting by \n), then pair.
    const items: PageItem[] = [];

    for (const entry of sortedEntries) {
      const entrySketch = sketchByEntry.get(entry.id);
      const displayDate = entryDisplayDate(entry);
      const displayTime = entry.time;
      const { left } = buildLeftPage(entry, displayDate, displayTime);

      const sketchPC: PageContent | undefined = entrySketch
        ? { type: 'sketch', content: entrySketch.dataUrl, sourceEntryId: entry.id, sourceDate: displayDate, sourceTime: displayTime }
        : undefined;

      items.push(...expandEntryToPages(entry, displayDate, displayTime, left, sketchPC));
    }

    // Pair page items into spreads
    let i = 0;
    while (i < items.length) {
      const item = items[i];

      if (item.sketchContent) {
        // This entry's last page has a sketch → sketch goes on the right
        spreads.push({
          entryId: item.entryId,
          date: item.date,
          time: item.time,
          left: item.content,
          right: item.sketchContent,
          isFirstSpread: true,
          rightHasTabs: false,
          rightTabDefault: 'sketch',
        });
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
          rightHasTabs: false,
          rightTabDefault: 'sketch',
        });
        i += nextItem ? 2 : 1;
      }
    }
  } else {
    // STT or AI mode: one spread per entry.
    // Literal \n is rendered as a regular newline (no page break).
    for (let i = 0; i < sortedEntries.length; i++) {
      const entry = sortedEntries[i];
      const entrySketch = sketchByEntry.get(entry.id);
      const hasSketch = !!entrySketch;
      const displayDate = entryDisplayDate(entry);
      const displayTime = entry.time;
      const { left, leftFallback } = buildLeftPage(entry, displayDate, displayTime);

      // Convert literal \n to real newlines in left page content
      const leftResolved: PageContent = left.type === 'text' && left.content
        ? { ...left, content: pageBreakToNewline(left.content) }
        : left;

      let right: PageContent;
      let rightAlt: PageContent | undefined;
      let rightHasTabs = false;

      const secondaryField = activeSection === 'stt' ? 'speechToText' : 'aiReport';
      let secondaryText = (entry[secondaryField] || '').trim();
      if (leftFallback === activeSection) {
        secondaryText = activeSection === 'stt'
          ? (entry.aiReport || '').trim()
          : (entry.speechToText || '').trim();
      }
      secondaryText = pageBreakToNewline(secondaryText);

      if (hasSketch) {
        right = {
          type: 'sketch',
          content: entrySketch!.dataUrl,
          sourceEntryId: entry.id,
          sourceDate: displayDate,
          sourceTime: displayTime,
        };
        rightAlt = { type: 'text', content: secondaryText, secondaryPage: 1, sourceEntryId: entry.id, sourceDate: displayDate, sourceTime: displayTime };
        rightHasTabs = true;
      } else {
        right = { type: 'text', content: secondaryText, secondaryPage: 1, sourceEntryId: entry.id, sourceDate: displayDate, sourceTime: displayTime };
      }

      spreads.push({
        entryId: entry.id,
        date: displayDate,
        time: displayTime,
        left: leftResolved,
        right,
        isFirstSpread: true,
        rightAltContent: rightAlt,
        rightHasTabs,
        rightTabDefault: 'sketch',
      });
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
  error?: string;
  appName: string;
}> {
  const res = await fetch('/api/entries');
  const data = await res.json();
  const appName = typeof data.appName === 'string' && data.appName.trim() ? data.appName.trim() : 'Daily Logger';
  return {
    entries: Array.isArray(data.entries) ? data.entries : [],
    sketches: Array.isArray(data.sketches) ? data.sketches : [],
    error: typeof data.error === 'string' ? data.error : undefined,
    appName,
  };
}

/* ──── main component ──── */

const JournalBook: React.FC = () => {
  const { t } = useReaderT();
  const [entries, setEntries] = useState<JournalEntry[]>([]);
  const [sketches, setSketches] = useState<PositionedSketch[]>([]);
  const [appTitle, setAppTitle] = useState('Daily Logger');
  const [loadError, setLoadError] = useState<string | null>(null);
  const [saveError, setSaveError] = useState<string | null>(null);
  const [currentPage, setCurrentPage] = useState(0);
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>(() => readStored(READER_SORT_KEY, ['asc', 'desc'], 'asc'));
  const [rightPageSetting, setRightPageSetting] = useState<RightPageSetting>(() => readStored(RIGHT_PAGE_KEY, ['none', 'ai', 'stt'], 'none'));
  const [activeSection, setActiveSection] = useState<JournalSection>(() => {
    const stored = readStored(RIGHT_PAGE_KEY, ['none', 'ai', 'stt'], 'none');
    return stored === 'none' ? 'journal' : stored;
  });
  const [rightTabOverride, setRightTabOverride] = useState<'sketch' | 'secondary' | null>(null);

  const [showSketchPlacer, setShowSketchPlacer] = useState(false);
  const [sketchingAfterEntryId, setSketchingAfterEntryId] = useState<string | null>(null);
  const [editingSketchId, setEditingSketchId] = useState<string | null>(null);

  const [editingEntryId, setEditingEntryId] = useState<string | null>(null);
  const [editedText, setEditedText] = useState('');
  const [editingLeftField, setEditingLeftField] = useState<'journal' | 'speechToText' | 'aiReport'>('journal');

  const [editingRightEntryId, setEditingRightEntryId] = useState<string | null>(null);
  const [editedRightText, setEditedRightText] = useState('');
  const [editingRightField, setEditingRightField] = useState<'journal' | 'speechToText' | 'aiReport'>('journal');

  const bookSpreadRef = useRef<HTMLDivElement>(null);

  /* ── load data ── */

  const reload = useCallback(async () => {
    try {
      const { entries: rows, sketches: sk, error, appName } = await fetchData();
      setEntries(rows);
      setSketches(sk);
      setAppTitle(appName);
      setLoadError(error ?? null);
    } catch {
      setLoadError(t('errLoadData'));
      setEntries([]);
      setSketches([]);
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

  const currentEntry = useMemo(() => {
    if (currentPage <= 0) return null;
    const sp = spreads[currentPage - 1];
    if (!sp) return null;
    return sortedEntries.find((e) => e.id === spreadPrimaryEntryId(sp)) ?? null;
  }, [currentPage, spreads, sortedEntries]);

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
      if (editingEntryId || editingRightEntryId) return;
      const key = e.key.toLowerCase();
      if (key === 'd' || key === 'arrowright') handleNext();
      if (key === 'a' || key === 'arrowleft') handlePrev();
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleNext, handlePrev, editingEntryId, editingRightEntryId]);

  /* ── actions ── */

  const handleAction = (action: JournalAction) => {
    if (action === 'sketch') {
      setShowSketchPlacer(true);
    } else {
      if (currentPage === 0) return;
      const spread = spreads[currentPage - 1];
      if (!spread) return;
      setSaveError(null);

      // Left page
      const leftEid = spread.left.sourceEntryId ?? spread.entryId;
      const leftEntry = entries.find((e) => e.id === leftEid);
      const leftFb = spread.left.leftFallbackSection;
      const leftField: 'journal' | 'speechToText' | 'aiReport' =
        leftFb === 'stt' ? 'speechToText' : leftFb === 'ai' ? 'aiReport' : 'journal';

      if (leftEntry) {
        setEditingEntryId(leftEid);
        setEditingLeftField(leftField);
        setEditedText(leftEntry[leftField] || '');
      }

      // Right page (text only, not sketch/empty)
      const rp = resolvedRight ?? spread.right;
      if (rp.type === 'text' && rp.sourceEntryId) {
        const rightEid = rp.sourceEntryId;
        const rightEntry = entries.find((e) => e.id === rightEid);
        if (rightEntry) {
          let rightField: 'journal' | 'speechToText' | 'aiReport' = 'journal';
          if (rp.secondaryPage !== undefined) {
            rightField = activeSection === 'stt' ? 'speechToText' : 'aiReport';
            if (spread.left.leftFallbackSection === activeSection) {
              rightField = activeSection === 'stt' ? 'aiReport' : 'speechToText';
            }
          }
          // Skip right editing when it's the same entry+field (page-break continuation)
          if (rightEid !== leftEid || rightField !== leftField) {
            setEditingRightEntryId(rightEid);
            setEditingRightField(rightField);
            setEditedRightText(rightEntry[rightField] || '');
          }
        }
      }
    }
  };

  const handleSaveText = async () => {
    if (!editingEntryId) return;
    setSaveError(null);
    try {
      // Save left page
      const leftBody: Record<string, string> = { id: editingEntryId };
      leftBody[editingLeftField] = editedText;
      const res = await fetch('/api/entry', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(leftBody),
      });
      const data = await res.json();
      if (!data.ok) { setSaveError(data.error || t('errSaveFailed')); return; }

      // Save right page if edited
      if (editingRightEntryId) {
        const rightBody: Record<string, string> = { id: editingRightEntryId };
        rightBody[editingRightField] = editedRightText;
        const res2 = await fetch('/api/entry', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(rightBody),
        });
        const data2 = await res2.json();
        if (!data2.ok) { setSaveError(data2.error || t('errSaveFailed')); return; }
      }

      await reload();
      setEditingEntryId(null);
      setEditingRightEntryId(null);
    } catch { setSaveError(t('errNetworkSave')); }
  };

  /* ── sketch CRUD ── */

  const handleCreateSketch = (afterEntryId: string) => {
    setSketchingAfterEntryId(afterEntryId);
    setEditingSketchId(null);
    setShowSketchPlacer(false);
  };

  const handleEditSketch = (sketchId: string) => {
    setEditingSketchId(sketchId);
    setSketchingAfterEntryId(null);
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
      if (existingId) {
        await handleDeleteSketch(existingId);
      }
      setSketchingAfterEntryId(null);
      setEditingSketchId(null);
      return;
    }

    try {
      const body: Record<string, unknown> = { dataUrl };
      if (existingId) {
        body.id = existingId;
      } else if (sketchingAfterEntryId) {
        body.afterEntryId = sketchingAfterEntryId;
      }
      const res = await fetch('/api/sketch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      const data = await res.json();
      if (!data.ok) { setSaveError(data.error || t('errSketchSave')); }
      await reload();
    } catch { setSaveError(t('errNetworkSketch')); }
    setSketchingAfterEntryId(null);
    setEditingSketchId(null);
  };

  const isDrawing = sketchingAfterEntryId !== null || editingSketchId !== null;
  const drawingInitialData = editingSketchId ? sketches.find((s) => s.id === editingSketchId)?.dataUrl : undefined;

  /* ── right page setting ── */

  const clearEditing = () => {
    setEditingEntryId(null);
    setEditedText('');
    setEditingRightEntryId(null);
    setEditedRightText('');
  };

  const handleRightPageSettingChange = (setting: RightPageSetting) => {
    clearEditing();
    setRightPageSetting(setting);
    persistStored(RIGHT_PAGE_KEY, setting);
    setActiveSection(setting === 'none' ? 'journal' : setting);
  };

  const handleBookmarkClick = (section: JournalSection) => {
    clearEditing();
    setActiveSection(section);
    const mapped: RightPageSetting = section === 'journal' ? 'none' : section;
    setRightPageSetting(mapped);
    persistStored(RIGHT_PAGE_KEY, mapped);
  };

  /* ── right tab toggle for spreads with tabs ── */

  useEffect(() => { setRightTabOverride(null); }, [currentPage]);

  const currentSpread = currentPage > 0 ? spreads[currentPage - 1] : undefined;
  const effectiveRightTab = currentSpread?.rightHasTabs
    ? (rightTabOverride ?? currentSpread.rightTabDefault ?? 'sketch')
    : null;

  const resolvedRight = useMemo(() => {
    if (!currentSpread) return undefined;
    if (effectiveRightTab === 'secondary' && currentSpread.rightAltContent) {
      return currentSpread.rightAltContent;
    }
    return currentSpread.right;
  }, [currentSpread, effectiveRightTab]);

  const resolvedSpread = useMemo(() => {
    if (!currentSpread) return undefined;
    if (resolvedRight && resolvedRight !== currentSpread.right) {
      return { ...currentSpread, right: resolvedRight };
    }
    return currentSpread;
  }, [currentSpread, resolvedRight]);

  const isEditing = !!editingEntryId || !!editingRightEntryId;

  return (
    <div className="flex h-dvh max-h-dvh min-h-0 flex-col overflow-hidden bg-[#2c1e14] text-[#d9c5b2] select-none font-sans">
      <div className="absolute top-6 left-8 text-[#d9c5b2] opacity-80 z-10">
        <h1 className="text-2xl tracking-widest uppercase font-light font-serif">
          {appTitle}{' '}
          <span className="text-xs opacity-50 block tracking-normal font-sans">{t('readerSubtitle')}</span>
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
        isEditTextOpen={isEditing}
        onSaveText={() => void handleSaveText()}
        rightPageSetting={rightPageSetting}
        onRightPageSettingChange={handleRightPageSettingChange}
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
                className="relative flex h-full min-h-0 w-full overflow-hidden rounded-lg bg-[#fdfaf2] shadow-2xl"
                initial={{ scale: 0.95, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                transition={{ duration: 0.5 }}
              >
                <div className="absolute inset-0 opacity-10 pointer-events-none bg-[url('https://www.transparenttextures.com/patterns/natural-paper.png')]" />

                {/* Click zones for page flip */}
                {!editingEntryId && (
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

                <Page
                  spread={resolvedSpread}
                  side="left"
                  rightPageSetting={rightPageSetting}
                  activeSection={activeSection}
                  editingEntryId={editingEntryId}
                  editText={editedText}
                  onTextChange={setEditedText}
                  rightTab={effectiveRightTab}
                  onRightTabChange={setRightTabOverride}
                />
                <Page
                  spread={resolvedSpread}
                  side="right"
                  rightPageSetting={rightPageSetting}
                  activeSection={activeSection}
                  editingEntryId={editingRightEntryId}
                  editText={editedRightText}
                  onTextChange={setEditedRightText}
                  rightTab={effectiveRightTab}
                  onRightTabChange={setRightTabOverride}
                  hasRightTabs={currentSpread?.rightHasTabs}
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
              />
              <BookmarkTab
                label={t('tabSpeech')}
                tabNumber="02"
                icon={<MessageSquare size={16} />}
                isActive={activeSection === 'stt'}
                onClick={() => handleBookmarkClick('stt')}
              />
              <BookmarkTab
                label={t('tabAi')}
                tabNumber="03"
                icon={<BrainCircuit size={16} />}
                isActive={activeSection === 'ai'}
                onClick={() => handleBookmarkClick('ai')}
              />
            </div>
          )}
        </div>
      </div>

      {/* Sketch placer modal */}
      {showSketchPlacer && (
        <SketchPlacer
          entries={sortedEntries}
          sketches={sketches}
          sortOrder={sortOrder}
          onCreateSketch={handleCreateSketch}
          onEditSketch={handleEditSketch}
          onDeleteSketch={handleDeleteSketch}
          onClose={() => setShowSketchPlacer(false)}
        />
      )}

      {/* Drawing canvas */}
      {isDrawing && (
        <DrawingCanvas
          onSave={handleSaveSketchCanvas}
          onClose={() => { setSketchingAfterEntryId(null); setEditingSketchId(null); }}
          initialData={drawingInitialData}
          sketchId={editingSketchId ?? undefined}
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
            <span className="text-sm uppercase tracking-widest font-sans">{t('footerPrev')}</span>
          </button>
          <div className="flex items-center space-x-3">
            <div className={cn('w-1 h-1 rounded-full', currentPage < 2 ? 'bg-white/50' : 'bg-white/20')} />
            <div className={cn('w-1.5 h-1.5 rounded-full', currentPage >= 2 && currentPage < spreadCount ? 'bg-white/50' : 'bg-white/20')} />
            <div className={cn('w-1 h-1 rounded-full', currentPage >= spreadCount ? 'bg-white/50' : 'bg-white/20')} />
          </div>
          <button
            type="button"
            onClick={handleNext}
            className="flex items-center space-x-2 group hover:text-[#d9c5b2] transition-colors"
          >
            <span className="text-sm uppercase tracking-widest font-sans">{t('footerNext')}</span>
            <ChevronRight className="w-5 h-5 group-hover:translate-x-1 transition-transform" />
          </button>
        </div>
        <div className="text-[10px] uppercase tracking-[0.3em] text-[#d9c5b2]/20 font-sans leading-relaxed">
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
  rightPageSetting: RightPageSetting;
  activeSection: JournalSection;
  editingEntryId: string | null;
  editText?: string;
  onTextChange?: (text: string) => void;
  rightTab?: 'sketch' | 'secondary' | null;
  onRightTabChange?: (tab: 'sketch' | 'secondary') => void;
  hasRightTabs?: boolean;
}> = ({ spread, side, rightPageSetting, activeSection, editingEntryId, editText, onTextChange, rightTab, onRightTabChange, hasRightTabs }) => {
  const { t } = useReaderT();
  if (!spread) {
    return (
      <div
        className={cn(
          'flex-1 h-full min-h-0 p-8 pr-10 flex items-center justify-center font-serif',
          side === 'right' && 'bg-[#fbf8ef] pl-10 pr-8',
        )}
      >
        <p className="text-zinc-400 italic tracking-widest uppercase text-xs opacity-40">{t('pageEmpty')}</p>
      </div>
    );
  }

  const pContent = side === 'left' ? spread.left : spread.right;

  /* ── sketch page ── */
  if (pContent.type === 'sketch') {
    const capDate = pContent.sourceDate ?? spread.date;
    return (
      <div className="flex-1 h-full min-h-0 bg-[#fbf8ef] relative group overflow-hidden flex flex-col">
        {/* Mini-tabs for right page */}
        {side === 'right' && hasRightTabs && (
          <RightPageTabs
            activeTab={rightTab ?? 'sketch'}
            onTabChange={onRightTabChange!}
            rightSetting={rightPageSetting}
            activeSection={activeSection}
          />
        )}
        <div className="flex-1 min-h-0 relative">
          <img
            src={pContent.content}
            alt={t('pageSketchAlt')}
            className="w-full h-full object-contain opacity-90 mix-blend-multiply group-hover:opacity-100 transition-opacity bg-[#fbf8ef]"
          />
          <div className="absolute inset-0 border-l border-black/5 pointer-events-none" />
          <div className="absolute bottom-12 right-12 text-black/10 text-[9px] font-sans font-bold uppercase tracking-[0.3em]">
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
        className={cn(
          'flex-1 h-full min-h-0 p-8 pr-10 flex items-center justify-center font-serif',
          side === 'right' && 'bg-[#fbf8ef] pl-10 pr-8 border-l border-black/5',
        )}
      >
        <p className="text-zinc-400 italic tracking-widest uppercase text-[10px] opacity-20">{t('pageBlank')}</p>
      </div>
    );
  }

  /* ── text page ── */
  const isRightSecondary = side === 'right' && pContent.secondaryPage !== undefined;
  const hasFallback = !!pContent.leftFallbackSection;
  const sectionTitle = isRightSecondary
    ? (rightPageSetting === 'stt' || (rightPageSetting === 'none' && activeSection === 'stt')
        ? t('pageVoiceTranscript') : t('pageIntelAnalysis'))
    : hasFallback
      ? (pContent.leftFallbackSection === 'stt' ? t('pageVoiceTranscript') : t('pageIntelAnalysis'))
      : t('pageDailyReflection');

  const bigHeaderTitle = hasFallback
    ? (pContent.leftFallbackSection === 'stt' ? t('pageVoiceTranscript') : t('pageIntelAnalysis'))
    : t('pageJournalEntry');

  const pageLabel = isRightSecondary ? pContent.secondaryPage : pContent.journalPage;
  const columnOwner = pContent.sourceEntryId ?? spread.entryId;
  const columnDate = pContent.sourceDate ?? spread.date;

  const showBigHeader = pContent.type === 'text' && !!pContent.showJournalEntryHeader;
  const showSubHeader = !showBigHeader && pContent.type === 'text' && pageLabel !== undefined;

  const showEditor =
    editingEntryId === columnOwner &&
    pContent.type === 'text';

  return (
    <div
      className={cn(
        'flex-1 h-full min-h-0 p-8 pr-10 flex flex-col relative overflow-hidden font-serif',
        side === 'right' ? 'pl-10 pr-8 bg-[#fbf8ef]' : 'bg-[#fdfaf2] border-r border-black/5',
      )}
    >
      {/* Mini-tabs for right page with secondary text */}
      {side === 'right' && hasRightTabs && (
        <RightPageTabs
          activeTab={rightTab ?? 'secondary'}
          onTabChange={onRightTabChange!}
          rightSetting={rightPageSetting}
          activeSection={activeSection}
        />
      )}

      {showBigHeader && (
        <div className="border-b border-black/5 pb-4 mb-6">
          <span className="text-xs uppercase tracking-widest text-black/40 font-sans">
            {columnDate} · {pContent.sourceTime ?? spread.time}
          </span>
          <h2 className="text-2xl font-light text-slate-800 leading-tight mt-1 font-serif">
            {bigHeaderTitle}
          </h2>
        </div>
      )}

      {showSubHeader && (
        <div className="border-b border-black/10 border-dashed pb-2 mb-6 flex justify-between items-end">
          <h3 className="text-sm font-light text-slate-400 font-serif italic tracking-wide">
            {pageLabel !== undefined && pageLabel > 1
              ? `${sectionTitle} ${t('pageContSuffix')}`
              : sectionTitle}
          </h3>
          <span className="text-[10px] text-black/20 font-sans">{columnDate}</span>
        </div>
      )}

      <div className="flex-1 text-slate-700 leading-relaxed text-base space-y-3 overflow-y-auto pr-2">
        {showEditor ? (
          <textarea
            value={editText}
            onChange={(e) => onTextChange?.(e.target.value)}
            className="w-full min-h-[60%] bg-transparent border border-black/10 rounded-md p-2 focus:ring-1 focus:ring-amber-700/30 resize-y font-serif text-slate-800 leading-relaxed"
            placeholder={t('pagePlaceholderEdit')}
            autoFocus
          />
        ) : (
          (pContent.content || '').split('\n').map((para, i) => <p key={i}>{para}</p>)
        )}
      </div>

      <div className="mt-8 flex justify-center shrink-0">
        <div className="w-16 h-1 bg-black/5 rounded-full" />
      </div>
    </div>
  );
};

/* ──── Right-page mini-tabs ──── */

const RightPageTabs: React.FC<{
  activeTab: 'sketch' | 'secondary';
  onTabChange: (tab: 'sketch' | 'secondary') => void;
  rightSetting: RightPageSetting;
  activeSection?: JournalSection;
}> = ({ activeTab, onTabChange, rightSetting, activeSection }) => {
  const { t } = useReaderT();
  const section = activeSection ?? (rightSetting === 'none' ? 'journal' : rightSetting);
  const secLabel = section === 'stt' ? t('rightTabStt') : t('rightTabAi');

  return (
    <div className="flex gap-1 mb-3 shrink-0 font-sans">
      <button
        onClick={() => onTabChange('sketch')}
        className={cn(
          'px-3 py-1 text-[10px] font-bold uppercase tracking-wider rounded-t-md transition-colors',
          activeTab === 'sketch'
            ? 'bg-[#e8dccb] text-[#5c4a36]'
            : 'bg-transparent text-black/30 hover:text-black/50 hover:bg-black/[0.03]',
        )}
      >
        {t('rightTabSketch')}
      </button>
      <button
        onClick={() => onTabChange('secondary')}
        className={cn(
          'px-3 py-1 text-[10px] font-bold uppercase tracking-wider rounded-t-md transition-colors',
          activeTab === 'secondary'
            ? 'bg-[#e8dccb] text-[#5c4a36]'
            : 'bg-transparent text-black/30 hover:text-black/50 hover:bg-black/[0.03]',
        )}
      >
        {secLabel}
      </button>
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
}> = ({ label, tabNumber, icon, isActive, onClick }) => {
  return (
    <motion.button
      type="button"
      onClick={onClick}
      className={cn(
        'px-6 py-3 rounded-r-md shadow-sm border-l-4 flex flex-col cursor-pointer transition-all w-36 items-start font-sans',
        isActive
          ? 'bg-[#e8dccb] text-[#5c4a36] border-[#8c7a66] translate-x-2'
          : 'bg-[#f4ead5]/80 text-[#8c7a66] border-[#8c7a66]/20 hover:bg-[#e8dccb] hover:translate-x-1',
      )}
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
