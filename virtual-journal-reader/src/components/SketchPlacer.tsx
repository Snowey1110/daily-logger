import React, { useState } from 'react';
import { X, Trash2, Plus, FileText, PenTool } from 'lucide-react';
import { cn, type JournalEntry, type PositionedSketch } from '../lib/utils';
import { useReaderT } from '../readerI18n';
import { useTheme } from './ThemeProvider';

interface SketchPlacerProps {
  entries: JournalEntry[];
  sketches: PositionedSketch[];
  sortOrder: 'asc' | 'desc';
  onCreateSketch: (afterEntryId: string) => void;
  onCreatePage: (date: string, time: string, afterEntryId: string) => void;
  onEditSketch: (sketchId: string) => void;
  onDeleteSketch: (sketchId: string) => void;
  onDeleteEntry: (entryId: string) => void;
  onClose: () => void;
}

type TimelineItem =
  | { kind: 'entry'; entry: JournalEntry }
  | { kind: 'sketch'; sketch: PositionedSketch; ownerEntryId: string };

function buildTimeline(
  entries: JournalEntry[],
  sketches: PositionedSketch[],
  sortOrder: 'asc' | 'desc',
): TimelineItem[] {
  const sorted = [...entries].sort((a, b) => {
    const ai = `${a.isoDate ?? a.date}|${String(a.rowIndex ?? 0).padStart(6, '0')}`;
    const bi = `${b.isoDate ?? b.date}|${String(b.rowIndex ?? 0).padStart(6, '0')}`;
    const cmp = ai.localeCompare(bi);
    return sortOrder === 'asc' ? cmp : -cmp;
  });

  const sketchesByAfter = new Map<string, PositionedSketch[]>();
  for (const sk of sketches) {
    const list = sketchesByAfter.get(sk.afterEntryId) ?? [];
    list.push(sk);
    sketchesByAfter.set(sk.afterEntryId, list);
  }

  const items: TimelineItem[] = [];
  for (const entry of sorted) {
    items.push({ kind: 'entry', entry });
    const entrySketchList = sketchesByAfter.get(entry.id) ?? [];
    for (const sk of entrySketchList) {
      items.push({ kind: 'sketch', sketch: sk, ownerEntryId: entry.id });
    }
  }
  return items;
}

function nearestEntryId(items: TimelineItem[], dividerIndex: number): string {
  for (let i = dividerIndex; i >= 0; i--) {
    if (items[i].kind === 'entry') return (items[i] as { kind: 'entry'; entry: JournalEntry }).entry.id;
  }
  if (items[0]?.kind === 'entry') return (items[0] as { kind: 'entry'; entry: JournalEntry }).entry.id;
  return '';
}

function nowDateStr(): string {
  const d = new Date();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${mm}/${dd}/${d.getFullYear()}`;
}

function nowTimeStr(): string {
  const d = new Date();
  let h = d.getHours();
  const ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12 || 12;
  const mm = String(d.getMinutes()).padStart(2, '0');
  return `${h}:${mm} ${ampm}`;
}

export const SketchPlacer: React.FC<SketchPlacerProps> = ({
  entries,
  sketches,
  sortOrder,
  onCreateSketch,
  onCreatePage,
  onEditSketch,
  onDeleteSketch,
  onDeleteEntry,
  onClose,
}) => {
  const { t } = useReaderT();
  const { bgTheme } = useTheme();
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);
  const [confirmDeleteEntryId, setConfirmDeleteEntryId] = useState<string | null>(null);
  const [expandedDivider, setExpandedDivider] = useState<number | null>(null);
  const [pageFormDivider, setPageFormDivider] = useState<number | null>(null);
  const [pageDate, setPageDate] = useState(nowDateStr);
  const [pageTime, setPageTime] = useState(nowTimeStr);

  const items = buildTimeline(entries, sketches, sortOrder);

  const handleDelete = (sketchId: string) => {
    if (confirmDeleteId === sketchId) {
      onDeleteSketch(sketchId);
      setConfirmDeleteId(null);
    } else {
      setConfirmDeleteId(sketchId);
    }
  };

  const handleDeleteEntry = (entryId: string) => {
    if (confirmDeleteEntryId === entryId) {
      onDeleteEntry(entryId);
      setConfirmDeleteEntryId(null);
    } else {
      setConfirmDeleteEntryId(entryId);
    }
  };

  const handleDividerClick = (dividerIdx: number) => {
    setExpandedDivider(expandedDivider === dividerIdx ? null : dividerIdx);
    setPageFormDivider(null);
  };

  const handleChooseSketch = (dividerIdx: number) => {
    const eid = nearestEntryId(items, dividerIdx);
    if (eid) onCreateSketch(eid);
  };

  const handleChoosePage = (dividerIdx: number) => {
    setPageFormDivider(dividerIdx);
    setPageDate(nowDateStr());
    setPageTime(nowTimeStr());
  };

  const handleSubmitPage = (dividerIdx: number) => {
    const eid = nearestEntryId(items, dividerIdx);
    if (eid && pageDate && pageTime) {
      onCreatePage(pageDate, pageTime, eid);
    }
  };

  return (
    <div className="fixed inset-0 z-[100] bg-black/80 flex items-center justify-center p-4 backdrop-blur-md font-sans">
      <div
        className="rounded-2xl shadow-[0_50px_100px_-20px_rgba(0,0,0,0.5)] w-full max-w-lg flex flex-col overflow-hidden max-h-[80vh] border"
        style={{
          backgroundColor: bgTheme.colors.bookInner,
          borderColor: bgTheme.colors.border,
        }}
      >
        <div
          className="p-4 flex items-center justify-between shrink-0"
          style={{ borderBottom: `1px solid ${bgTheme.colors.border}` }}
        >
          <h3
            className="font-semibold uppercase tracking-widest text-xs"
            style={{ color: bgTheme.colors.text }}
          >
            {t('sketchPlacerTitle')}
          </h3>
          <button
            onClick={onClose}
            className="p-2 hover:bg-black/5 rounded-full transition-colors"
            style={{ color: bgTheme.colors.textMuted }}
          >
            <X size={20} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-4 space-y-0">
          {items.map((item, idx) => {
            const isLast = idx === items.length - 1;
            const dividerIdx = idx;

            return (
              <div key={item.kind === 'entry' ? `e-${item.entry.id}` : `s-${item.sketch.id}`} className={item.kind === 'entry' ? 'group/entry' : undefined}>
                {item.kind === 'entry' ? (
                  <div
                    className="py-3 px-3 text-sm font-medium font-serif flex items-center justify-between"
                    style={{ color: bgTheme.colors.text }}
                  >
                    <span>{item.entry.date} {item.entry.time}</span>
                    {confirmDeleteEntryId === item.entry.id ? (
                      <div className="flex items-center gap-1 shrink-0">
                        <button
                          onClick={() => handleDeleteEntry(item.entry.id)}
                          className="px-2 py-1 text-xs font-semibold text-red-600 bg-red-50 hover:bg-red-100 rounded transition-colors"
                        >
                          {t('deleteEntry')}
                        </button>
                        <button
                          onClick={() => setConfirmDeleteEntryId(null)}
                          className="px-2 py-1 text-xs font-semibold rounded transition-colors"
                          style={{ color: bgTheme.colors.textMuted }}
                        >
                          {t('sketchCancelBtn')}
                        </button>
                      </div>
                    ) : (
                      <button
                        onClick={() => handleDeleteEntry(item.entry.id)}
                        className="p-1.5 hover:bg-red-50 rounded hover:text-red-500 transition-colors shrink-0 opacity-0 group-hover/entry:opacity-100"
                        style={{ color: bgTheme.colors.textMuted }}
                        title={t('deleteEntry')}
                      >
                        <Trash2 size={14} />
                      </button>
                    )}
                  </div>
                ) : (
                  <div className="py-2 px-3">
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => onEditSketch(item.sketch.id)}
                        className="h-10 w-full rounded-md border overflow-hidden hover:opacity-80 transition-opacity cursor-pointer"
                        style={{ borderColor: bgTheme.colors.border, backgroundColor: bgTheme.colors.bookInner }}
                      >
                        <img
                          src={item.sketch.dataUrl}
                          alt=""
                          className="h-full w-full object-contain"
                        />
                      </button>
                      {confirmDeleteId === item.sketch.id ? (
                        <div className="flex items-center gap-1 shrink-0">
                          <button
                            onClick={() => handleDelete(item.sketch.id)}
                            className="px-2 py-1 text-xs font-semibold text-red-600 bg-red-50 hover:bg-red-100 rounded transition-colors"
                          >
                            {t('sketchDeleteBtn')}
                          </button>
                          <button
                            onClick={() => setConfirmDeleteId(null)}
                            className="px-2 py-1 text-xs font-semibold rounded transition-colors"
                            style={{ color: bgTheme.colors.textMuted, backgroundColor: `${bgTheme.colors.textMuted}20` }}
                          >
                            {t('sketchCancelBtn')}
                          </button>
                        </div>
                      ) : (
                        <button
                          onClick={() => handleDelete(item.sketch.id)}
                          className="p-1.5 hover:bg-red-50 rounded hover:text-red-500 transition-colors shrink-0"
                          style={{ color: bgTheme.colors.textMuted }}
                          title={t('sketchDeleteBtn')}
                        >
                          <Trash2 size={16} />
                        </button>
                      )}
                    </div>
                  </div>
                )}

                {/* Divider / insert controls */}
                {!isLast && (
                  <div className="px-3">
                    {expandedDivider === dividerIdx ? (
                      pageFormDivider === dividerIdx ? (
                        /* Inline date/time form */
                        <div
                          className="flex flex-wrap items-center gap-2 py-2 px-3 rounded-lg border"
                          style={{ borderColor: bgTheme.colors.border, backgroundColor: `${bgTheme.colors.textMuted}10` }}
                        >
                          <label className="text-[10px] uppercase tracking-wider font-bold" style={{ color: bgTheme.colors.textMuted }}>
                            {t('insertPageDateLabel')}
                          </label>
                          <input
                            type="text"
                            value={pageDate}
                            onChange={(e) => setPageDate(e.target.value)}
                            className="w-28 px-2 py-1 text-xs rounded border bg-transparent focus:outline-none"
                            style={{ borderColor: bgTheme.colors.border, color: bgTheme.colors.text }}
                            placeholder="MM/DD/YYYY"
                          />
                          <label className="text-[10px] uppercase tracking-wider font-bold" style={{ color: bgTheme.colors.textMuted }}>
                            {t('insertPageTimeLabel')}
                          </label>
                          <input
                            type="text"
                            value={pageTime}
                            onChange={(e) => setPageTime(e.target.value)}
                            className="w-24 px-2 py-1 text-xs rounded border bg-transparent focus:outline-none"
                            style={{ borderColor: bgTheme.colors.border, color: bgTheme.colors.text }}
                            placeholder="HH:MM AM"
                          />
                          <button
                            onClick={() => handleSubmitPage(dividerIdx)}
                            className="px-3 py-1 text-xs font-semibold rounded transition-colors text-white"
                            style={{ backgroundColor: bgTheme.colors.tabs.journal.active }}
                          >
                            {t('insertPageCreate')}
                          </button>
                          <button
                            onClick={() => { setPageFormDivider(null); setExpandedDivider(null); }}
                            className="px-2 py-1 text-xs font-semibold rounded transition-colors"
                            style={{ color: bgTheme.colors.textMuted }}
                          >
                            {t('sketchCancelBtn')}
                          </button>
                        </div>
                      ) : (
                        /* Choice: New Page / New Sketch */
                        <div className="flex items-center justify-center gap-3 py-2">
                          <button
                            onClick={() => handleChoosePage(dividerIdx)}
                            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-md border transition-colors hover:opacity-80"
                            style={{ borderColor: bgTheme.colors.border, color: bgTheme.colors.text }}
                          >
                            <FileText size={14} />
                            {t('insertChoicePage')}
                          </button>
                          <button
                            onClick={() => handleChooseSketch(dividerIdx)}
                            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-md border transition-colors hover:opacity-80"
                            style={{ borderColor: bgTheme.colors.border, color: bgTheme.colors.text }}
                          >
                            <PenTool size={14} />
                            {t('insertChoiceSketch')}
                          </button>
                          <button
                            onClick={() => setExpandedDivider(null)}
                            className="px-2 py-1 text-xs rounded transition-colors"
                            style={{ color: bgTheme.colors.textMuted }}
                          >
                            {t('sketchCancelBtn')}
                          </button>
                        </div>
                      )
                    ) : (
                      /* Collapsed divider with + icon */
                      <button
                        onClick={() => handleDividerClick(dividerIdx)}
                        className="group w-full flex items-center gap-3 py-1.5 transition-colors rounded-md hover:bg-black/[0.03]"
                        title={t('sketchInsertHere')}
                      >
                        <div className="flex-1 h-px transition-colors" style={{ backgroundColor: bgTheme.colors.border }} />
                        <Plus
                          size={14}
                          className="shrink-0 transition-colors"
                          style={{ color: bgTheme.colors.textMuted }}
                        />
                        <div className="flex-1 h-px transition-colors" style={{ backgroundColor: bgTheme.colors.border }} />
                      </button>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};
