import React, { useState } from 'react';
import { X, Trash2, Plus } from 'lucide-react';
import { cn, type JournalEntry, type PositionedSketch } from '../lib/utils';
import { useReaderT } from '../readerI18n';

interface SketchPlacerProps {
  entries: JournalEntry[];
  sketches: PositionedSketch[];
  sortOrder: 'asc' | 'desc';
  onCreateSketch: (afterEntryId: string) => void;
  onEditSketch: (sketchId: string) => void;
  onDeleteSketch: (sketchId: string) => void;
  onClose: () => void;
}

export const SketchPlacer: React.FC<SketchPlacerProps> = ({
  entries,
  sketches,
  sortOrder,
  onCreateSketch,
  onEditSketch,
  onDeleteSketch,
  onClose,
}) => {
  const { t } = useReaderT();
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);

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

  const formatEntryLabel = (e: JournalEntry) => {
    return `${e.date} ${e.time}`;
  };

  const handleDelete = (sketchId: string) => {
    if (confirmDeleteId === sketchId) {
      onDeleteSketch(sketchId);
      setConfirmDeleteId(null);
    } else {
      setConfirmDeleteId(sketchId);
    }
  };

  return (
    <div className="fixed inset-0 z-[100] bg-black/80 flex items-center justify-center p-4 backdrop-blur-md font-sans">
      <div className="bg-[#fdfaf2] rounded-2xl shadow-[0_50px_100px_-20px_rgba(0,0,0,0.5)] w-full max-w-lg flex flex-col overflow-hidden max-h-[80vh] border border-[#d9c5b2]/20">
        <div className="p-4 border-b border-[#d9c5b2]/10 flex items-center justify-between bg-[#fbf8ef] shrink-0">
          <h3 className="font-semibold text-slate-800 uppercase tracking-widest text-xs">
            {t('sketchPlacerTitle')}
          </h3>
          <button
            onClick={onClose}
            className="p-2 hover:bg-black/5 rounded-full text-slate-400 transition-colors"
          >
            <X size={20} />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-4 space-y-0">
          {sorted.map((entry, idx) => {
            const entrySketchList = sketchesByAfter.get(entry.id) ?? [];
            const isLast = idx === sorted.length - 1;

            return (
              <div key={entry.id}>
                {/* Entry label */}
                <div className="py-3 px-3 text-sm font-medium text-slate-700 font-serif">
                  {formatEntryLabel(entry)}
                </div>

                {/* Sketches placed after this entry */}
                {entrySketchList.map((sk) => (
                  <div key={sk.id} className="py-2 px-3">
                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => onEditSketch(sk.id)}
                        className="h-10 w-full rounded-md border border-black/10 overflow-hidden hover:border-black/30 transition-colors cursor-pointer bg-white"
                      >
                        <img
                          src={sk.dataUrl}
                          alt=""
                          className="h-full w-full object-contain mix-blend-multiply"
                        />
                      </button>
                      {confirmDeleteId === sk.id ? (
                        <div className="flex items-center gap-1 shrink-0">
                          <button
                            onClick={() => handleDelete(sk.id)}
                            className="px-2 py-1 text-xs font-semibold text-red-600 bg-red-50 hover:bg-red-100 rounded transition-colors"
                          >
                            {t('sketchDeleteBtn')}
                          </button>
                          <button
                            onClick={() => setConfirmDeleteId(null)}
                            className="px-2 py-1 text-xs font-semibold text-slate-500 bg-slate-100 hover:bg-slate-200 rounded transition-colors"
                          >
                            {t('sketchCancelBtn')}
                          </button>
                        </div>
                      ) : (
                        <button
                          onClick={() => handleDelete(sk.id)}
                          className="p-1.5 hover:bg-red-50 rounded text-slate-400 hover:text-red-500 transition-colors shrink-0"
                          title={t('sketchDeleteBtn')}
                        >
                          <Trash2 size={16} />
                        </button>
                      )}
                    </div>
                  </div>
                ))}

                {/* Divider / insert button */}
                {!isLast && (
                  <button
                    onClick={() => onCreateSketch(entry.id)}
                    className={cn(
                      'group w-full flex items-center gap-3 py-1.5 px-3 transition-colors rounded-md',
                      'hover:bg-black/[0.03]'
                    )}
                    title={t('sketchInsertHere')}
                  >
                    <div className="flex-1 h-px bg-black/10 group-hover:bg-amber-700/40 transition-colors" />
                    <Plus
                      size={14}
                      className="text-black/15 group-hover:text-amber-700/60 transition-colors shrink-0"
                    />
                    <div className="flex-1 h-px bg-black/10 group-hover:bg-amber-700/40 transition-colors" />
                  </button>
                )}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};
