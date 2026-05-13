import React, { useRef, useEffect, useState, useCallback } from 'react';
import {
  X, Eraser, Upload, GripVertical, Trash2,
  Type, PenTool, ImageIcon, ChevronUp, ChevronDown,
} from 'lucide-react';
import { useReaderT } from '../readerI18n';
import { useTheme } from './ThemeProvider';
import type { PageImage } from '../lib/utils';

type LayerKind = 'text' | 'sketch' | 'images';

interface CompositePageEditorProps {
  entryId: string;
  entryDate: string;
  entryTime: string;
  pageWidth: number;
  pageHeight: number;
  initialText: string;
  initialSketchDataUrl?: string;
  initialImages: PageImage[];
  initialLayerOrder: LayerKind[];
  defaultLayer: LayerKind;
  onSave: (text: string, sketchDataUrl: string, images: PageImage[], layerOrder: LayerKind[]) => void;
  onClose: () => void;
}

let _imgCounter = 0;
function nextImgId(): string {
  return `img_${Date.now()}_${++_imgCounter}`;
}

export const CompositePageEditor: React.FC<CompositePageEditorProps> = ({
  entryDate,
  entryTime,
  pageWidth,
  pageHeight,
  initialText,
  initialSketchDataUrl,
  initialImages,
  initialLayerOrder,
  defaultLayer,
  onSave,
  onClose,
}) => {
  const { t } = useReaderT();
  const { bgTheme } = useTheme();

  const [text, setText] = useState(initialText);
  const [images, setImages] = useState<PageImage[]>(initialImages);
  const [layerOrder, setLayerOrder] = useState<LayerKind[]>(
    initialLayerOrder.length === 3 ? initialLayerOrder : ['text', 'sketch', 'images'],
  );
  const [activeLayer, setActiveLayer] = useState<LayerKind>(defaultLayer);

  /* ── sketch canvas state ── */
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const ctxRef = useRef<CanvasRenderingContext2D | null>(null);
  const [isDrawing, setIsDrawing] = useState(false);
  const [color, setColor] = useState('#000000');
  const [lineWidth, setLineWidth] = useState(3);
  const [isErasing, setIsErasing] = useState(false);
  const [eraserSize, setEraserSize] = useState(20);
  const pageAreaRef = useRef<HTMLDivElement>(null);
  const sketchLoaded = useRef(false);

  /* ── image interaction state ── */
  const [selectedImgId, setSelectedImgId] = useState<string | null>(null);
  const [dragging, setDragging] = useState<{ id: string; startX: number; startY: number; origX: number; origY: number } | null>(null);
  const [resizing, setResizing] = useState<{ id: string; startX: number; startY: number; origW: number; origH: number; origImgX: number; origImgY: number } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  /* ── init canvas (resize only, no initial image loading) ── */
  const initCanvas = useCallback(() => {
    const canvas = canvasRef.current;
    const pageArea = pageAreaRef.current;
    if (!canvas || !pageArea) return;

    const { width, height } = pageArea.getBoundingClientRect();
    if (width === 0 || height === 0) return;

    const prevW = canvas.width;
    const prevH = canvas.height;

    if (prevW === width && prevH === height) return;

    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = prevW;
    tempCanvas.height = prevH;
    const tempCtx = tempCanvas.getContext('2d');
    if (tempCtx && prevW > 0 && prevH > 0) {
      tempCtx.drawImage(canvas, 0, 0);
    }

    canvas.width = width;
    canvas.height = height;

    const ctx = canvas.getContext('2d');
    if (ctx) {
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.strokeStyle = isErasing ? '#000000' : color;
      ctx.lineWidth = isErasing ? eraserSize : lineWidth;
      ctx.globalCompositeOperation = isErasing ? 'destination-out' : 'source-over';
      ctxRef.current = ctx;

      if (prevW > 0 && prevH > 0) {
        ctx.drawImage(tempCanvas, 0, 0, prevW, prevH, 0, 0, width, height);
      }
    }
  }, [color, lineWidth, isErasing, eraserSize]);

  useEffect(() => {
    const canvas = canvasRef.current;
    const pageArea = pageAreaRef.current;
    if (!canvas || !pageArea) return;

    const { width, height } = pageArea.getBoundingClientRect();
    if (width === 0 || height === 0) return;
    canvas.width = width;
    canvas.height = height;

    const ctx = canvas.getContext('2d');
    if (ctx) {
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.strokeStyle = color;
      ctx.lineWidth = lineWidth;
      ctx.globalCompositeOperation = 'source-over';
      ctxRef.current = ctx;
    }

    if (initialSketchDataUrl && !sketchLoaded.current) {
      const img = new Image();
      img.onload = () => {
        if (ctxRef.current && canvas.width > 0) {
          ctxRef.current.drawImage(img, 0, 0, canvas.width, canvas.height);
        }
        sketchLoaded.current = true;
      };
      img.src = initialSketchDataUrl;
    }
    sketchLoaded.current = true;

    window.addEventListener('resize', initCanvas);
    return () => window.removeEventListener('resize', initCanvas);
  }, []); // mount only

  useEffect(() => {
    if (ctxRef.current) {
      ctxRef.current.strokeStyle = isErasing ? '#000000' : color;
      ctxRef.current.lineWidth = isErasing ? eraserSize : lineWidth;
      ctxRef.current.globalCompositeOperation = isErasing ? 'destination-out' : 'source-over';
    }
  }, [color, lineWidth, isErasing, eraserSize]);

  /* ── drawing helpers ── */
  const getCanvasCoords = (e: React.MouseEvent | React.TouchEvent) => {
    const canvas = canvasRef.current;
    if (!canvas) return { x: 0, y: 0 };
    const rect = canvas.getBoundingClientRect();
    let cx: number, cy: number;
    if ('touches' in e) {
      cx = e.touches[0].clientX;
      cy = e.touches[0].clientY;
    } else {
      cx = e.clientX;
      cy = e.clientY;
    }
    return { x: cx - rect.left, y: cy - rect.top };
  };

  const startDrawing = (e: React.MouseEvent | React.TouchEvent) => {
    if (activeLayer !== 'sketch') return;
    if ('touches' in e) e.preventDefault();
    const { x, y } = getCanvasCoords(e);
    ctxRef.current?.beginPath();
    ctxRef.current?.moveTo(x, y);
    setIsDrawing(true);
  };

  const draw = (e: React.MouseEvent | React.TouchEvent) => {
    if (!isDrawing || activeLayer !== 'sketch') return;
    if ('touches' in e) e.preventDefault();
    const { x, y } = getCanvasCoords(e);
    ctxRef.current?.lineTo(x, y);
    ctxRef.current?.stroke();
  };

  const stopDrawing = () => {
    if (isDrawing) {
      ctxRef.current?.closePath();
      setIsDrawing(false);
    }
  };

  const clearCanvas = () => {
    const canvas = canvasRef.current;
    if (canvas && ctxRef.current) {
      const prevOp = ctxRef.current.globalCompositeOperation;
      ctxRef.current.globalCompositeOperation = 'source-over';
      ctxRef.current.clearRect(0, 0, canvas.width, canvas.height);
      ctxRef.current.globalCompositeOperation = prevOp;
    }
  };

  /* ── image helpers ── */
  const addImageFromFile = (file: File) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result as string;
      setImages((prev) => [
        ...prev,
        { id: nextImgId(), dataUrl, x: 0.25, y: 0.25, width: 0.5, height: 0.4 },
      ]);
    };
    reader.readAsDataURL(file);
  };

  const handleUpload = () => fileInputRef.current?.click();

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) addImageFromFile(file);
    e.target.value = '';
  };

  useEffect(() => {
    const handler = (e: ClipboardEvent) => {
      const items = e.clipboardData?.items;
      if (!items) return;
      for (const item of items) {
        if (item.type.startsWith('image/')) {
          const file = item.getAsFile();
          if (file) {
            e.preventDefault();
            addImageFromFile(file);
          }
          break;
        }
      }
    };
    window.addEventListener('paste', handler);
    return () => window.removeEventListener('paste', handler);
  }, []);

  const deleteImage = (id: string) => {
    setImages((prev) => prev.filter((img) => img.id !== id));
    if (selectedImgId === id) setSelectedImgId(null);
  };

  /* ── image drag / resize (mouse) ── */
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      const pageArea = pageAreaRef.current;
      if (!pageArea) return;
      const rect = pageArea.getBoundingClientRect();

      if (dragging) {
        const dx = (e.clientX - dragging.startX) / rect.width;
        const dy = (e.clientY - dragging.startY) / rect.height;
        setImages((prev) =>
          prev.map((img) =>
            img.id === dragging.id
              ? { ...img, x: Math.max(0, Math.min(1, dragging.origX + dx)), y: Math.max(0, Math.min(1, dragging.origY + dy)) }
              : img,
          ),
        );
      }

      if (resizing) {
        const dx = (e.clientX - resizing.startX) / rect.width;
        const dy = (e.clientY - resizing.startY) / rect.height;
        setImages((prev) =>
          prev.map((img) =>
            img.id === resizing.id
              ? {
                  ...img,
                  width: Math.max(0.05, resizing.origW + dx),
                  height: Math.max(0.05, resizing.origH + dy),
                }
              : img,
          ),
        );
      }
    };
    const handleMouseUp = () => {
      setDragging(null);
      setResizing(null);
    };
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [dragging, resizing]);

  const handleImageMouseDown = (e: React.MouseEvent, img: PageImage) => {
    if (activeLayer !== 'images') return;
    e.stopPropagation();
    e.preventDefault();
    setSelectedImgId(img.id);
    setDragging({ id: img.id, startX: e.clientX, startY: e.clientY, origX: img.x, origY: img.y });
  };

  const handleResizeMouseDown = (e: React.MouseEvent, img: PageImage) => {
    e.stopPropagation();
    setResizing({ id: img.id, startX: e.clientX, startY: e.clientY, origW: img.width, origH: img.height, origImgX: img.x, origImgY: img.y });
  };

  /* ── layer reorder ── */
  const moveLayer = (kind: LayerKind, dir: -1 | 1) => {
    setLayerOrder((prev) => {
      const idx = prev.indexOf(kind);
      const newIdx = idx + dir;
      if (newIdx < 0 || newIdx >= prev.length) return prev;
      const next = [...prev];
      [next[idx], next[newIdx]] = [next[newIdx], next[idx]];
      return next;
    });
  };

  /* ── drag reorder via pointer ── */
  const [dragLayerIdx, setDragLayerIdx] = useState<number | null>(null);
  const [dragOverIdx, setDragOverIdx] = useState<number | null>(null);

  const handleLayerDragStart = (idx: number) => {
    setDragLayerIdx(idx);
  };
  const handleLayerDragOver = (e: React.DragEvent, idx: number) => {
    e.preventDefault();
    setDragOverIdx(idx);
  };
  const handleLayerDrop = (idx: number) => {
    if (dragLayerIdx !== null && dragLayerIdx !== idx) {
      setLayerOrder((prev) => {
        const next = [...prev];
        const [moved] = next.splice(dragLayerIdx, 1);
        next.splice(idx, 0, moved);
        return next;
      });
    }
    setDragLayerIdx(null);
    setDragOverIdx(null);
  };

  /* ── save ── */
  const handleSave = () => {
    const canvas = canvasRef.current;
    let sketchDataUrl = '';
    if (canvas) {
      const ctx = canvas.getContext('2d', { willReadFrequently: true });
      if (ctx) {
        const { data } = ctx.getImageData(0, 0, canvas.width, canvas.height);
        let blank = true;
        for (let i = 3; i < data.length; i += 4) {
          if (data[i] !== 0) { blank = false; break; }
        }
        if (!blank) sketchDataUrl = canvas.toDataURL('image/png');
      }
    }
    onSave(text, sketchDataUrl, images, layerOrder);
  };

  /* ── render helpers ── */
  const layerIcon = (kind: LayerKind) => {
    switch (kind) {
      case 'text': return <Type size={14} />;
      case 'sketch': return <PenTool size={14} />;
      case 'images': return <ImageIcon size={14} />;
    }
  };

  const layerLabel = (kind: LayerKind) => {
    switch (kind) {
      case 'text': return t('layerText');
      case 'sketch': return t('layerSketch');
      case 'images': return t('layerImages');
    }
  };

  const zIndex = (kind: LayerKind) => layerOrder.indexOf(kind) + 1;

  const saveBtnBg = bgTheme.cover.isDark ? '#4f46e5' : '#334155';

  const editorTotalWidth = pageWidth + 192;

  return (
    <div className="fixed inset-0 z-[100] bg-black/80 flex flex-col items-center justify-center backdrop-blur-md font-sans">
      {/* Top bar */}
      <div
        className="flex items-center justify-between px-6 py-3 rounded-t-2xl"
        style={{ backgroundColor: bgTheme.colors.bookInner, borderBottom: `1px solid ${bgTheme.colors.border}`, width: editorTotalWidth }}
      >
        <h3 className="font-semibold uppercase tracking-widest text-xs" style={{ color: bgTheme.colors.text }}>
          {t('compositeEditorTitle')}
        </h3>
        <div className="flex items-center gap-3">
          {activeLayer === 'sketch' && (
            <div className="flex items-center gap-2 border-r pr-3 mr-1" style={{ borderColor: bgTheme.colors.border }}>
              {/* Eraser toggle */}
              <button
                onClick={() => setIsErasing((v) => !v)}
                className="p-1.5 rounded-full transition-colors border"
                style={{
                  backgroundColor: isErasing ? '#ef4444' : 'transparent',
                  borderColor: isErasing ? '#ef4444' : bgTheme.colors.border,
                  color: isErasing ? '#fff' : bgTheme.colors.textMuted,
                }}
                title={isErasing ? 'Draw mode' : 'Eraser mode'}
              >
                <Eraser size={16} />
              </button>
              {isErasing ? (
                /* Eraser size slider */
                <input
                  type="range"
                  min="5"
                  max="50"
                  value={eraserSize}
                  onChange={(e) => setEraserSize(parseInt(e.target.value))}
                  className="w-20 accent-red-500"
                />
              ) : (
                /* Drawing tools */
                <>
                  <input
                    type="color"
                    value={color}
                    onChange={(e) => setColor(e.target.value)}
                    className="w-6 h-6 rounded-full cursor-pointer border-0 p-0 overflow-hidden"
                  />
                  <input
                    type="range"
                    min="1"
                    max="20"
                    value={lineWidth}
                    onChange={(e) => setLineWidth(parseInt(e.target.value))}
                    className="w-20 accent-slate-600"
                  />
                </>
              )}
              {/* Clear sketch button */}
              <button
                onClick={clearCanvas}
                className="p-1.5 hover:bg-black/5 rounded-full transition-colors"
                style={{ color: bgTheme.colors.textMuted }}
                title={t('clearCanvas')}
              >
                <Trash2 size={16} />
              </button>
            </div>
          )}
          {activeLayer === 'images' && (
            <div className="flex items-center gap-2 border-r pr-3 mr-1" style={{ borderColor: bgTheme.colors.border }}>
              <button
                onClick={handleUpload}
                className="flex items-center gap-1 px-3 py-1.5 text-xs font-semibold rounded-md border transition-colors hover:opacity-80"
                style={{ borderColor: bgTheme.colors.border, color: bgTheme.colors.text }}
              >
                <Upload size={14} />
                {t('imageUpload')}
              </button>
              <span className="text-[10px] opacity-50" style={{ color: bgTheme.colors.textMuted }}>
                {t('imagePaste')}
              </span>
              <input ref={fileInputRef} type="file" accept="image/*" className="hidden" onChange={handleFileChange} />
            </div>
          )}
          <button
            onClick={handleSave}
            className="px-5 py-2 rounded-full text-white font-semibold text-xs tracking-widest uppercase shadow-lg hover:opacity-90 transition-opacity"
            style={{ backgroundColor: saveBtnBg }}
          >
            {t('compositeEditorSave')}
          </button>
          <button onClick={onClose} className="p-2 hover:bg-black/5 rounded-full transition-colors" style={{ color: bgTheme.colors.textMuted }}>
            <X size={20} />
          </button>
        </div>
      </div>

      {/* Main area: page + layer panel */}
      <div className="flex gap-0 min-h-0 overflow-hidden" style={{ width: editorTotalWidth, height: pageHeight }}>
        {/* Page area – fixed to rendered page dimensions */}
        <div
          ref={pageAreaRef}
          className="relative p-8 pr-10 flex flex-col font-serif shrink-0 overflow-hidden"
          style={{ backgroundColor: bgTheme.colors.bookInner, width: pageWidth, height: pageHeight }}
          onMouseDown={(e) => {
            if (activeLayer === 'images' && e.target === e.currentTarget) setSelectedImgId(null);
          }}
        >
          {/* Header – matches rendered Page's showBigHeader block, pointer-events:none so canvas receives draw events */}
          <div className="pb-4 mb-6 relative" style={{ borderBottom: `1px solid ${bgTheme.colors.border}`, pointerEvents: 'none' }}>
            <span className="text-xs uppercase tracking-widest font-sans opacity-40" style={{ color: bgTheme.colors.textMuted }}>
              {entryDate} · {entryTime}
            </span>
            <h2 className="text-2xl font-light leading-tight mt-1 font-serif" style={{ color: bgTheme.colors.text }}>
              {t('pageJournalEntry')}
            </h2>
          </div>

          {/* Content wrapper – matches rendered Page's flex-1 content area */}
          <div
            className="flex-1 relative min-h-0"
            style={{ zIndex: zIndex('text'), pointerEvents: activeLayer === 'text' ? 'auto' : 'none' }}
          >
            <textarea
              value={text}
              onChange={(e) => setText(e.target.value)}
              className="absolute inset-0 w-full h-full bg-transparent resize-none font-serif text-base leading-relaxed focus:outline-none overflow-auto pr-2"
              style={{ color: bgTheme.colors.text }}
              placeholder={activeLayer === 'text' ? 'Type here...' : ''}
            />
          </div>

          {/* Bottom divider – matches rendered Page */}
          <div className="mt-8 flex justify-center shrink-0" style={{ pointerEvents: 'none' }}>
            <div className="w-16 h-1 rounded-full" style={{ backgroundColor: bgTheme.colors.border }} />
          </div>

          {/* Sketch layer – covers full page area (inset-0) to match overlay in rendered Page */}
          <canvas
            ref={canvasRef}
            className="absolute inset-0 touch-none"
            style={{
              zIndex: zIndex('sketch'),
              pointerEvents: activeLayer === 'sketch' ? 'auto' : 'none',
              cursor: activeLayer === 'sketch' ? (isErasing ? 'cell' : 'crosshair') : 'default',
              width: '100%',
              height: '100%',
            }}
            onMouseDown={startDrawing}
            onMouseMove={draw}
            onMouseUp={stopDrawing}
            onMouseOut={stopDrawing}
            onTouchStart={startDrawing}
            onTouchMove={draw}
            onTouchEnd={stopDrawing}
          />

          {/* Image layer – covers full page area (inset-0) to match overlay in rendered Page */}
          <div
            className="absolute inset-0"
            style={{ zIndex: zIndex('images'), pointerEvents: activeLayer === 'images' ? 'auto' : 'none' }}
            onMouseDown={(e) => {
              if (e.target === e.currentTarget) setSelectedImgId(null);
            }}
          >
            {images.map((img) => {
              const isSelected = selectedImgId === img.id;
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
                    cursor: activeLayer === 'images' ? 'move' : 'default',
                  }}
                  onMouseDown={(e) => handleImageMouseDown(e, img)}
                >
                  <img src={img.dataUrl} alt="" className="w-full h-full object-contain select-none pointer-events-none" draggable={false} />
                  {isSelected && activeLayer === 'images' && (
                    <>
                      <button
                        className="absolute -top-3 -right-3 w-6 h-6 bg-red-500 text-white rounded-full flex items-center justify-center shadow hover:bg-red-600 transition-colors"
                        onClick={(e) => { e.stopPropagation(); deleteImage(img.id); }}
                        title={t('imageDelete')}
                      >
                        <X size={14} />
                      </button>
                      <div
                        className="absolute -bottom-2 -right-2 w-5 h-5 bg-blue-500 rounded-sm cursor-se-resize flex items-center justify-center shadow"
                        onMouseDown={(e) => { e.stopPropagation(); handleResizeMouseDown(e, img); }}
                      >
                        <GripVertical size={10} className="text-white rotate-45" />
                      </div>
                    </>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* Layer panel (Photoshop-style) */}
        <div
          className="w-48 flex flex-col border-l shrink-0"
          style={{ backgroundColor: bgTheme.colors.bg, borderColor: bgTheme.colors.border }}
        >
          <div
            className="px-3 py-2 text-[10px] font-bold uppercase tracking-widest"
            style={{ color: bgTheme.colors.textMuted, borderBottom: `1px solid ${bgTheme.colors.border}` }}
          >
            Layers
          </div>
          <div className="flex-1 overflow-y-auto">
            {[...layerOrder].reverse().map((kind, displayIdx) => {
              const realIdx = layerOrder.length - 1 - displayIdx;
              const isActive = activeLayer === kind;
              return (
                <div
                  key={kind}
                  draggable
                  onDragStart={() => handleLayerDragStart(realIdx)}
                  onDragOver={(e) => handleLayerDragOver(e, realIdx)}
                  onDrop={() => handleLayerDrop(realIdx)}
                  onClick={() => setActiveLayer(kind)}
                  className="flex items-center gap-2 px-3 py-2.5 cursor-pointer transition-colors border-b select-none"
                  style={{
                    backgroundColor: isActive ? `${bgTheme.colors.tabs.journal.active}20` : 'transparent',
                    borderColor: bgTheme.colors.border,
                    borderLeft: isActive ? `3px solid ${bgTheme.colors.tabs.journal.active}` : '3px solid transparent',
                    opacity: dragOverIdx === realIdx ? 0.5 : 1,
                  }}
                >
                  <GripVertical size={12} className="cursor-grab opacity-40" style={{ color: bgTheme.colors.textMuted }} />
                  <span style={{ color: isActive ? bgTheme.colors.tabs.journal.active : bgTheme.colors.textMuted }}>
                    {layerIcon(kind)}
                  </span>
                  <span
                    className="text-xs font-semibold tracking-wide"
                    style={{ color: isActive ? bgTheme.colors.text : bgTheme.colors.textMuted }}
                  >
                    {layerLabel(kind)}
                  </span>
                  <div className="ml-auto flex items-center gap-0.5">
                    <button
                      onClick={(e) => { e.stopPropagation(); moveLayer(kind, 1); }}
                      className="p-0.5 rounded hover:bg-black/5 transition-colors"
                      style={{ color: bgTheme.colors.textMuted }}
                      title={t('layerMoveUp')}
                    >
                      <ChevronUp size={12} />
                    </button>
                    <button
                      onClick={(e) => { e.stopPropagation(); moveLayer(kind, -1); }}
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
        </div>
      </div>

      {/* Bottom rounded corners */}
      <div className="h-4 rounded-b-2xl" style={{ backgroundColor: bgTheme.colors.bookInner, width: editorTotalWidth }} />
    </div>
  );
};
