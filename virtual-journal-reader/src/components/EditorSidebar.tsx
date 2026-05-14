import React from 'react';
import {
  X, Eraser, Upload, GripVertical, Trash2,
  Type, PenTool, ImageIcon, ChevronUp, ChevronDown, Save,
} from 'lucide-react';
import { useReaderT } from '../readerI18n';
import { useTheme } from './ThemeProvider';

export type LayerKind = 'text' | 'sketch' | 'images';

interface EditorSidebarProps {
  activeLayer: LayerKind;
  onLayerChange: (layer: LayerKind) => void;
  layerOrder: LayerKind[];
  onMoveLayer: (kind: LayerKind, dir: -1 | 1) => void;
  onReorderLayers: (fromIdx: number, toIdx: number) => void;
  /* sketch tools */
  color: string;
  lineWidth: number;
  isErasing: boolean;
  eraserSize: number;
  onColorChange: (c: string) => void;
  onLineWidthChange: (w: number) => void;
  onSetErasing: (v: boolean) => void;
  onEraserSizeChange: (s: number) => void;
  onClearCanvas: () => void;
  /* image tools */
  onUploadImage: () => void;
  /* actions */
  onSave: () => void;
  onClose: () => void;
  /* display info */
  entryDate: string;
  entryTime: string;
  isMobile?: boolean;
}

const LAYER_ICON: Record<LayerKind, React.ReactNode> = {
  text: <Type size={14} />,
  sketch: <PenTool size={14} />,
  images: <ImageIcon size={14} />,
};

export const EditorSidebar: React.FC<EditorSidebarProps> = ({
  activeLayer, onLayerChange, layerOrder, onMoveLayer, onReorderLayers,
  color, lineWidth, isErasing, eraserSize,
  onColorChange, onLineWidthChange, onSetErasing, onEraserSizeChange, onClearCanvas,
  onUploadImage, onSave, onClose, entryDate, entryTime, isMobile,
}) => {
  const { t } = useReaderT();
  const { bgTheme } = useTheme();
  const [expanded, setExpanded] = React.useState(false);

  const layerLabel = (kind: LayerKind) => {
    switch (kind) {
      case 'text': return t('layerText');
      case 'sketch': return t('layerSketch');
      case 'images': return t('layerImages');
    }
  };

  const [dragIdx, setDragIdx] = React.useState<number | null>(null);
  const [overIdx, setOverIdx] = React.useState<number | null>(null);

  const saveBtnBg = bgTheme.cover.isDark ? '#4f46e5' : '#334155';

  if (isMobile) {
    return (
      <div
        className="fixed left-0 right-0 bottom-0 z-[110] flex flex-col shadow-2xl font-sans"
        style={{
          backgroundColor: bgTheme.colors.bookInner,
          borderTop: `1px solid ${bgTheme.colors.border}`,
          maxHeight: expanded ? '60vh' : 'auto',
          transition: 'max-height 0.3s ease',
        }}
      >
        {/* Compact toolbar row */}
        <div className="flex items-center gap-1 px-2 py-1.5 shrink-0">
          {/* Layer tabs */}
          {layerOrder.map((kind) => (
            <button
              key={kind}
              onClick={() => { onLayerChange(kind); setExpanded(true); }}
              className="flex items-center gap-1 px-3 py-2 rounded-lg text-xs font-semibold transition-colors min-h-[40px]"
              style={{
                backgroundColor: activeLayer === kind ? `${bgTheme.colors.tabs.journal.active}20` : 'transparent',
                color: activeLayer === kind ? bgTheme.colors.text : bgTheme.colors.textMuted,
                borderBottom: activeLayer === kind ? `2px solid ${bgTheme.colors.tabs.journal.active}` : '2px solid transparent',
              }}
            >
              {LAYER_ICON[kind]}
              {layerLabel(kind)}
            </button>
          ))}
          <div className="flex-1" />
          <button
            onClick={onSave}
            className="flex items-center gap-1 px-3 py-2 rounded-lg text-white font-semibold text-xs min-h-[40px]"
            style={{ backgroundColor: saveBtnBg }}
          >
            <Save size={14} />
            {t('compositeEditorSave')}
          </button>
          <button onClick={onClose} className="p-2 rounded-full min-h-[40px] min-w-[40px] flex items-center justify-center" style={{ color: bgTheme.colors.textMuted }}>
            <X size={18} />
          </button>
        </div>

        {/* Expandable tools area */}
        {expanded && (
          <div className="px-3 py-2 overflow-y-auto" style={{ borderTop: `1px solid ${bgTheme.colors.border}` }}>
            {activeLayer === 'sketch' && (
              <div className="flex flex-wrap items-center gap-3">
                <button
                  onClick={() => onSetErasing(false)}
                  className="flex items-center gap-1 px-3 py-2 rounded-lg border text-xs font-semibold min-h-[40px]"
                  style={{
                    backgroundColor: !isErasing ? `${bgTheme.colors.tabs.journal.active}20` : 'transparent',
                    borderColor: !isErasing ? bgTheme.colors.tabs.journal.active : bgTheme.colors.border,
                    color: !isErasing ? bgTheme.colors.text : bgTheme.colors.textMuted,
                  }}
                >
                  <PenTool size={14} /> Draw
                </button>
                <button
                  onClick={() => onSetErasing(true)}
                  className="flex items-center gap-1 px-3 py-2 rounded-lg border text-xs font-semibold min-h-[40px]"
                  style={{
                    backgroundColor: isErasing ? '#ef444420' : 'transparent',
                    borderColor: isErasing ? '#ef4444' : bgTheme.colors.border,
                    color: isErasing ? '#ef4444' : bgTheme.colors.textMuted,
                  }}
                >
                  <Eraser size={14} /> Eraser
                </button>
                {isErasing ? (
                  <input type="range" min="5" max="50" value={eraserSize} onChange={(e) => onEraserSizeChange(parseInt(e.target.value))} className="w-24 accent-red-500" />
                ) : (
                  <>
                    <input type="color" value={color} onChange={(e) => onColorChange(e.target.value)} className="w-8 h-8 rounded-full cursor-pointer border-0 p-0" />
                    <input type="range" min="1" max="20" value={lineWidth} onChange={(e) => onLineWidthChange(parseInt(e.target.value))} className="w-24 accent-slate-600" />
                  </>
                )}
                <button onClick={onClearCanvas} className="flex items-center gap-1 px-2 py-2 text-xs rounded-lg border min-h-[40px]" style={{ borderColor: bgTheme.colors.border, color: bgTheme.colors.textMuted }}>
                  <Trash2 size={14} /> {t('clearCanvas')}
                </button>
              </div>
            )}
            {activeLayer === 'images' && (
              <div className="flex items-center gap-3">
                <button onClick={onUploadImage} className="flex items-center gap-2 px-3 py-2 text-xs font-semibold rounded-lg border min-h-[40px]" style={{ borderColor: bgTheme.colors.border, color: bgTheme.colors.text }}>
                  <Upload size={14} /> {t('imageUpload')}
                </button>
                <span className="text-[10px] opacity-50" style={{ color: bgTheme.colors.textMuted }}>{t('imagePaste')}</span>
              </div>
            )}
            {activeLayer === 'text' && (
              <p className="text-[10px] opacity-50" style={{ color: bgTheme.colors.textMuted }}>
                Tap on the page to edit text directly.
              </p>
            )}
          </div>
        )}
      </div>
    );
  }

  return (
    <div
      className="fixed left-0 top-0 bottom-0 w-56 z-[110] flex flex-col shadow-2xl font-sans"
      style={{ backgroundColor: bgTheme.colors.bookInner, borderRight: `1px solid ${bgTheme.colors.border}` }}
    >
      {/* Header */}
      <div
        className="flex items-center justify-between px-4 py-3 shrink-0"
        style={{ borderBottom: `1px solid ${bgTheme.colors.border}` }}
      >
        <h3 className="font-semibold uppercase tracking-widest text-[10px]" style={{ color: bgTheme.colors.text }}>
          {t('compositeEditorTitle')}
        </h3>
        <button onClick={onClose} className="p-1 hover:bg-black/5 rounded-full transition-colors" style={{ color: bgTheme.colors.textMuted }}>
          <X size={16} />
        </button>
      </div>

      {/* Entry info */}
      <div className="px-4 py-2 shrink-0" style={{ borderBottom: `1px solid ${bgTheme.colors.border}` }}>
        <span className="text-[10px] uppercase tracking-widest font-sans opacity-50" style={{ color: bgTheme.colors.textMuted }}>
          {entryDate} · {entryTime}
        </span>
      </div>

      {/* Save button */}
      <div className="px-4 py-3 shrink-0" style={{ borderBottom: `1px solid ${bgTheme.colors.border}` }}>
        <button
          onClick={onSave}
          className="w-full flex items-center justify-center gap-2 py-2.5 rounded-lg text-white font-semibold text-xs tracking-widest uppercase shadow-lg hover:opacity-90 transition-opacity"
          style={{ backgroundColor: saveBtnBg }}
        >
          <Save size={14} />
          {t('compositeEditorSave')}
        </button>
      </div>

      {/* Layers */}
      <div className="shrink-0" style={{ borderBottom: `1px solid ${bgTheme.colors.border}` }}>
        <div className="px-4 py-2 text-[10px] font-bold uppercase tracking-widest" style={{ color: bgTheme.colors.textMuted }}>
          Layers
        </div>
        {[...layerOrder].reverse().map((kind, displayIdx) => {
          const realIdx = layerOrder.length - 1 - displayIdx;
          const isActive = activeLayer === kind;
          return (
            <div
              key={kind}
              draggable
              onDragStart={() => setDragIdx(realIdx)}
              onDragOver={(e) => { e.preventDefault(); setOverIdx(realIdx); }}
              onDrop={() => { if (dragIdx !== null && dragIdx !== realIdx) onReorderLayers(dragIdx, realIdx); setDragIdx(null); setOverIdx(null); }}
              onClick={() => onLayerChange(kind)}
              className="flex items-center gap-2 px-4 py-2.5 cursor-pointer transition-colors select-none"
              style={{
                backgroundColor: isActive ? `${bgTheme.colors.tabs.journal.active}20` : 'transparent',
                borderLeft: isActive ? `3px solid ${bgTheme.colors.tabs.journal.active}` : '3px solid transparent',
                opacity: overIdx === realIdx ? 0.5 : 1,
              }}
            >
              <GripVertical size={12} className="cursor-grab opacity-40" style={{ color: bgTheme.colors.textMuted }} />
              <span style={{ color: isActive ? bgTheme.colors.tabs.journal.active : bgTheme.colors.textMuted }}>
                {LAYER_ICON[kind]}
              </span>
              <span
                className="text-xs font-semibold tracking-wide flex-1"
                style={{ color: isActive ? bgTheme.colors.text : bgTheme.colors.textMuted }}
              >
                {layerLabel(kind)}
              </span>
              <div className="flex items-center gap-0.5">
                <button
                  onClick={(e) => { e.stopPropagation(); onMoveLayer(kind, 1); }}
                  className="p-0.5 rounded hover:bg-black/5 transition-colors"
                  style={{ color: bgTheme.colors.textMuted }}
                  title={t('layerMoveUp')}
                >
                  <ChevronUp size={12} />
                </button>
                <button
                  onClick={(e) => { e.stopPropagation(); onMoveLayer(kind, -1); }}
                  className="p-0.5 rounded hover:bg-black/5 transition-colors"
                  style={{ color: bgTheme.colors.textMuted }}
                  title={t('layerMoveDown')}
                >
                  <ChevronDown size={12} />
                </button>
              </div>
            </div>
          );
        })}
      </div>

      {/* Active-layer tools */}
      <div className="flex-1 overflow-y-auto px-4 py-3">
        <div className="text-[10px] font-bold uppercase tracking-widest mb-3" style={{ color: bgTheme.colors.textMuted }}>
          Tools
        </div>

        {activeLayer === 'sketch' && (
          <div className="flex flex-col gap-3">
            {/* Three tool buttons */}
            <div className="flex flex-col gap-1.5">
              {/* Draw */}
              <button
                onClick={() => onSetErasing(false)}
                className="flex items-center gap-2 px-3 py-2 rounded-lg transition-colors border"
                style={{
                  backgroundColor: !isErasing ? `${bgTheme.colors.tabs.journal.active}20` : 'transparent',
                  borderColor: !isErasing ? bgTheme.colors.tabs.journal.active : bgTheme.colors.border,
                  color: !isErasing ? bgTheme.colors.text : bgTheme.colors.textMuted,
                }}
              >
                <PenTool size={16} />
                <span className="text-xs font-semibold">Draw</span>
              </button>
              {/* Eraser */}
              <button
                onClick={() => onSetErasing(true)}
                className="flex items-center gap-2 px-3 py-2 rounded-lg transition-colors border"
                style={{
                  backgroundColor: isErasing ? '#ef444420' : 'transparent',
                  borderColor: isErasing ? '#ef4444' : bgTheme.colors.border,
                  color: isErasing ? '#ef4444' : bgTheme.colors.textMuted,
                }}
              >
                <Eraser size={16} />
                <span className="text-xs font-semibold">Eraser</span>
              </button>
            </div>

            {/* Context-specific controls */}
            {isErasing ? (
              <div>
                <label className="text-[10px] uppercase tracking-wider mb-1 block" style={{ color: bgTheme.colors.textMuted }}>
                  Size
                </label>
                <input
                  type="range"
                  min="5" max="50"
                  value={eraserSize}
                  onChange={(e) => onEraserSizeChange(parseInt(e.target.value))}
                  className="w-full accent-red-500"
                />
              </div>
            ) : (
              <>
                <div className="flex items-center gap-2">
                  <label className="text-[10px] uppercase tracking-wider" style={{ color: bgTheme.colors.textMuted }}>
                    Color
                  </label>
                  <input
                    type="color"
                    value={color}
                    onChange={(e) => onColorChange(e.target.value)}
                    className="w-7 h-7 rounded-full cursor-pointer border-0 p-0 overflow-hidden"
                  />
                </div>
                <div>
                  <label className="text-[10px] uppercase tracking-wider mb-1 block" style={{ color: bgTheme.colors.textMuted }}>
                    Width
                  </label>
                  <input
                    type="range"
                    min="1" max="20"
                    value={lineWidth}
                    onChange={(e) => onLineWidthChange(parseInt(e.target.value))}
                    className="w-full accent-slate-600"
                  />
                </div>
              </>
            )}

            <button
              onClick={onClearCanvas}
              className="flex items-center gap-2 px-3 py-2 text-xs font-semibold rounded-lg border transition-colors hover:opacity-80 mt-1"
              style={{ borderColor: bgTheme.colors.border, color: bgTheme.colors.textMuted }}
            >
              <Trash2 size={14} />
              {t('clearCanvas')}
            </button>
          </div>
        )}

        {activeLayer === 'images' && (
          <div className="flex flex-col gap-3">
            <button
              onClick={onUploadImage}
              className="flex items-center gap-2 px-3 py-2 text-xs font-semibold rounded-lg border transition-colors hover:opacity-80"
              style={{ borderColor: bgTheme.colors.border, color: bgTheme.colors.text }}
            >
              <Upload size={14} />
              {t('imageUpload')}
            </button>
            <span className="text-[10px] opacity-50" style={{ color: bgTheme.colors.textMuted }}>
              {t('imagePaste')}
            </span>
          </div>
        )}

        {activeLayer === 'text' && (
          <p className="text-[10px] opacity-50" style={{ color: bgTheme.colors.textMuted }}>
            Click on the page to edit text directly.
          </p>
        )}
      </div>
    </div>
  );
};
