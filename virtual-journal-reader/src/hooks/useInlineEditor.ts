import React, { useRef, useState, useEffect, useCallback } from 'react';
import type { PageImage } from '../lib/utils';
import {
  captureCanvasSnapshot,
  createCanvasHistory,
  isShiftStraightLine,
  pushUndoSnapshot,
  redoCanvas,
  restoreCanvasSnapshot,
  straightLineTargetWithGuideSnap,
  undoCanvas,
} from '../lib/sketchDrawing';

export type LayerKind = 'text' | 'sketch' | 'images';

let _imgCounter = 0;
function nextImgId(): string {
  return `img_${Date.now()}_${++_imgCounter}`;
}

function buildEraserCursor(size: number): string {
  const r = Math.max(size / 2, 3);
  const svgSize = Math.ceil(r * 2 + 4);
  const center = svgSize / 2;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${svgSize}" height="${svgSize}"><circle cx="${center}" cy="${center}" r="${r}" fill="none" stroke="rgba(239,68,68,0.7)" stroke-width="2"/><circle cx="${center}" cy="${center}" r="1.5" fill="rgba(239,68,68,0.9)"/></svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}") ${Math.round(center)} ${Math.round(center)}, cell`;
}

export interface InlineEditorState {
  text: string;
  setText: (t: string) => void;
  images: PageImage[];
  layerOrder: LayerKind[];
  activeLayer: LayerKind;
  setActiveLayer: (l: LayerKind) => void;
  /* sketch */
  canvasRef: React.RefObject<HTMLCanvasElement | null>;
  color: string;
  lineWidth: number;
  isErasing: boolean;
  eraserSize: number;
  eraserCursor: string;
  setColor: (c: string) => void;
  setLineWidth: (w: number) => void;
  setIsErasing: (v: boolean) => void;
  toggleErasing: () => void;
  setEraserSize: (s: number) => void;
  clearCanvas: () => void;
  startDrawing: (e: React.MouseEvent | React.TouchEvent) => void;
  draw: (e: React.MouseEvent | React.TouchEvent) => void;
  stopDrawing: () => void;
  /* images */
  selectedImgId: string | null;
  setSelectedImgId: (id: string | null) => void;
  handleImageMouseDown: (e: React.MouseEvent, img: PageImage) => void;
  handleResizeMouseDown: (e: React.MouseEvent, img: PageImage) => void;
  deleteImage: (id: string) => void;
  handleUpload: () => void;
  handleFileChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  fileInputRef: React.RefObject<HTMLInputElement | null>;
  /* layers */
  moveLayer: (kind: LayerKind, dir: -1 | 1) => void;
  reorderLayers: (fromIdx: number, toIdx: number) => void;
  zIndex: (kind: LayerKind) => number;
  /* save */
  getSketchDataUrl: () => string;
  getSavePayload: () => { text: string; sketchDataUrl: string; images: PageImage[]; layerOrder: LayerKind[] };
  /* init */
  initCanvasForElement: (el: HTMLElement) => void;
}

export function useInlineEditor(opts: {
  initialText: string;
  initialSketchDataUrl?: string;
  initialImages: PageImage[];
  initialLayerOrder: LayerKind[];
  defaultLayer: LayerKind;
  pageAreaRef: React.RefObject<HTMLElement | null>;
}): InlineEditorState {
  const { initialText, initialSketchDataUrl, initialImages, initialLayerOrder, defaultLayer, pageAreaRef } = opts;

  const [text, setText] = useState(initialText);
  const [images, setImages] = useState<PageImage[]>(initialImages);
  const [layerOrder, setLayerOrder] = useState<LayerKind[]>(
    initialLayerOrder.length === 3 ? initialLayerOrder : ['text', 'sketch', 'images'],
  );
  const [activeLayer, setActiveLayer] = useState<LayerKind>(defaultLayer);

  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const ctxRef = useRef<CanvasRenderingContext2D | null>(null);
  const [isDrawingStroke, setIsDrawingStroke] = useState(false);
  const [color, setColor] = useState('#000000');
  const [lineWidth, setLineWidth] = useState(3);
  const [isErasing, setIsErasing] = useState(false);
  const [eraserSize, setEraserSize] = useState(20);
  const sketchLoaded = useRef(false);
  const strokeStartRef = useRef<{ x: number; y: number } | null>(null);
  const strokeSnapshotRef = useRef<ImageData | null>(null);
  const historyRef = useRef(createCanvasHistory());

  const eraserCursor = isErasing ? buildEraserCursor(eraserSize) : 'crosshair';

  const [selectedImgId, setSelectedImgId] = useState<string | null>(null);
  const [dragging, setDragging] = useState<{ id: string; startX: number; startY: number; origX: number; origY: number } | null>(null);
  const [resizing, setResizing] = useState<{ id: string; startX: number; startY: number; origW: number; origH: number } | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  /* ── canvas sizing: CSS layout must match page; bitmap attrs alone shrink the hit box ── */
  const applyCtxDefaults = useCallback((ctx: CanvasRenderingContext2D) => {
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = isErasing ? '#000000' : color;
    ctx.lineWidth = isErasing ? eraserSize : lineWidth;
    ctx.globalCompositeOperation = isErasing ? 'destination-out' : 'source-over';
  }, [color, lineWidth, isErasing, eraserSize]);

  const syncCanvasSize = useCallback(() => {
    const canvas = canvasRef.current;
    const pageEl = pageAreaRef.current;
    if (!canvas || !pageEl) return;

    const rect = pageEl.getBoundingClientRect();
    const width = Math.round(rect.width);
    const height = Math.round(rect.height);
    if (width === 0 || height === 0) return;

    const prevW = canvas.width;
    const prevH = canvas.height;
    const layoutUnchanged =
      prevW === width &&
      prevH === height &&
      canvas.style.width === `${width}px` &&
      canvas.style.height === `${height}px`;

    if (layoutUnchanged) {
      if (!ctxRef.current) {
        const ctx = canvas.getContext('2d');
        if (ctx) {
          applyCtxDefaults(ctx);
          ctxRef.current = ctx;
        }
      }
      return;
    }

    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = prevW;
    tempCanvas.height = prevH;
    const tempCtx = tempCanvas.getContext('2d');
    if (tempCtx && prevW > 0 && prevH > 0) tempCtx.drawImage(canvas, 0, 0);

    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;
    canvas.style.position = 'absolute';
    canvas.style.left = '0';
    canvas.style.top = '0';

    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    applyCtxDefaults(ctx);
    ctxRef.current = ctx;

    if (prevW > 0 && prevH > 0) {
      ctx.drawImage(tempCanvas, 0, 0, prevW, prevH, 0, 0, width, height);
    } else if (initialSketchDataUrl && !sketchLoaded.current) {
      const img = new Image();
      img.onload = () => {
        if (ctxRef.current && canvas.width > 0) {
          ctxRef.current.drawImage(img, 0, 0, canvas.width, canvas.height);
          sketchLoaded.current = true;
        }
      };
      img.src = initialSketchDataUrl;
    }
  }, [applyCtxDefaults, initialSketchDataUrl, pageAreaRef]);

  const initCanvasForElement = useCallback((_el: HTMLElement) => {
    syncCanvasSize();
  }, [syncCanvasSize]);

  useEffect(() => {
    const canvas = canvasRef.current;
    const pageEl = pageAreaRef.current;
    if (!canvas) return;

    const runSync = () => syncCanvasSize();
    runSync();
    const raf1 = requestAnimationFrame(() => {
      runSync();
      requestAnimationFrame(runSync);
    });

    const ro = new ResizeObserver(runSync);
    ro.observe(canvas);
    if (pageEl) ro.observe(pageEl);

    window.addEventListener('resize', runSync);
    return () => {
      cancelAnimationFrame(raf1);
      ro.disconnect();
      window.removeEventListener('resize', runSync);
    };
  }, [syncCanvasSize, pageAreaRef]);

  useEffect(() => {
    if (ctxRef.current) applyCtxDefaults(ctxRef.current);
  }, [applyCtxDefaults]);

  useEffect(() => {
    if (activeLayer === 'sketch') {
      requestAnimationFrame(() => syncCanvasSize());
    }
  }, [activeLayer, syncCanvasSize]);

  /* ── drawing ── */
  const getCanvasCoords = (e: React.MouseEvent | React.TouchEvent) => {
    const canvas = canvasRef.current;
    if (!canvas) return { x: 0, y: 0 };
    const rect = canvas.getBoundingClientRect();
    if (rect.width === 0 || rect.height === 0) return { x: 0, y: 0 };
    let cx: number, cy: number;
    if ('touches' in e) { cx = e.touches[0].clientX; cy = e.touches[0].clientY; }
    else { cx = e.clientX; cy = e.clientY; }
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    return { x: (cx - rect.left) * scaleX, y: (cy - rect.top) * scaleY };
  };

  const startDrawing = (e: React.MouseEvent | React.TouchEvent) => {
    if (activeLayer !== 'sketch') return;
    if ('touches' in e) e.preventDefault();
    syncCanvasSize();
    const { x, y } = getCanvasCoords(e);
    strokeStartRef.current = { x, y };
    const ctx = ctxRef.current;
    const canvas = canvasRef.current;
    if (ctx && canvas) {
      strokeSnapshotRef.current = captureCanvasSnapshot(ctx, canvas);
      pushUndoSnapshot(historyRef.current, strokeSnapshotRef.current);
    }
    ctx?.beginPath();
    ctx?.moveTo(x, y);
    setIsDrawingStroke(true);
  };

  const draw = (e: React.MouseEvent | React.TouchEvent) => {
    if (!isDrawingStroke || activeLayer !== 'sketch') return;
    if ('touches' in e) e.preventDefault();
    const { x, y } = getCanvasCoords(e);
    const ctx = ctxRef.current;
    if (!ctx) return;

    const start = strokeStartRef.current;
    if (
      isShiftStraightLine(e) &&
      !isErasing &&
      start &&
      strokeSnapshotRef.current
    ) {
      const canvas = canvasRef.current;
      restoreCanvasSnapshot(ctx, strokeSnapshotRef.current, canvas ?? undefined);
      const end = straightLineTargetWithGuideSnap(start, { x, y });
      ctx.beginPath();
      ctx.moveTo(start.x, start.y);
      ctx.lineTo(end.x, end.y);
      ctx.stroke();
    } else {
      ctx.lineTo(x, y);
      ctx.stroke();
    }
  };

  const stopDrawing = () => {
    if (isDrawingStroke) {
      ctxRef.current?.closePath();
      setIsDrawingStroke(false);
      strokeStartRef.current = null;
      strokeSnapshotRef.current = null;
    }
  };

  const clearCanvas = () => {
    const canvas = canvasRef.current;
    if (canvas && ctxRef.current) {
      const prevOp = ctxRef.current.globalCompositeOperation;
      pushUndoSnapshot(historyRef.current, captureCanvasSnapshot(ctxRef.current, canvas));
      ctxRef.current.globalCompositeOperation = 'source-over';
      ctxRef.current.clearRect(0, 0, canvas.width, canvas.height);
      ctxRef.current.globalCompositeOperation = prevOp;
    }
  };

  const undoSketch = useCallback(() => {
    const canvas = canvasRef.current;
    const ctx = ctxRef.current;
    if (!canvas || !ctx) return;
    if (undoCanvas(ctx, canvas, historyRef.current)) applyCtxDefaults(ctx);
  }, [applyCtxDefaults]);

  const redoSketch = useCallback(() => {
    const canvas = canvasRef.current;
    const ctx = ctxRef.current;
    if (!canvas || !ctx) return;
    if (redoCanvas(ctx, canvas, historyRef.current)) applyCtxDefaults(ctx);
  }, [applyCtxDefaults]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (activeLayer !== 'sketch') return;
      if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
      const key = e.key.toLowerCase();
      if (key === 'z' && !e.shiftKey) {
        e.preventDefault();
        undoSketch();
      } else if (key === 'y' || (key === 'z' && e.shiftKey)) {
        e.preventDefault();
        redoSketch();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [activeLayer, undoSketch, redoSketch]);

  /* ── images ── */
  const addImageFromFile = useCallback((file: File) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result as string;
      setImages((prev) => [...prev, { id: nextImgId(), dataUrl, x: 0.25, y: 0.25, width: 0.5, height: 0.4 }]);
    };
    reader.readAsDataURL(file);
  }, []);

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
          if (file) { e.preventDefault(); addImageFromFile(file); }
          break;
        }
      }
    };
    window.addEventListener('paste', handler);
    return () => window.removeEventListener('paste', handler);
  }, [addImageFromFile]);

  const deleteImage = (id: string) => {
    setImages((prev) => prev.filter((img) => img.id !== id));
    if (selectedImgId === id) setSelectedImgId(null);
  };

  /* ── image drag / resize ── */
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      const el = pageAreaRef.current;
      if (!el) return;
      const rect = el.getBoundingClientRect();
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
              ? { ...img, width: Math.max(0.05, resizing.origW + dx), height: Math.max(0.05, resizing.origH + dy) }
              : img,
          ),
        );
      }
    };
    const handleMouseUp = () => { setDragging(null); setResizing(null); };
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => { window.removeEventListener('mousemove', handleMouseMove); window.removeEventListener('mouseup', handleMouseUp); };
  }, [dragging, resizing, pageAreaRef]);

  const handleImageMouseDown = (e: React.MouseEvent, img: PageImage) => {
    if (activeLayer !== 'images') return;
    e.stopPropagation();
    e.preventDefault();
    setSelectedImgId(img.id);
    setDragging({ id: img.id, startX: e.clientX, startY: e.clientY, origX: img.x, origY: img.y });
  };

  const handleResizeMouseDown = (e: React.MouseEvent, img: PageImage) => {
    e.stopPropagation();
    setResizing({ id: img.id, startX: e.clientX, startY: e.clientY, origW: img.width, origH: img.height });
  };

  /* ── layer ordering ── */
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

  const reorderLayers = (fromIdx: number, toIdx: number) => {
    setLayerOrder((prev) => {
      const next = [...prev];
      const [moved] = next.splice(fromIdx, 1);
      next.splice(toIdx, 0, moved);
      return next;
    });
  };

  const zIndex = (kind: LayerKind) => layerOrder.indexOf(kind) + 1;

  /* ── save ── */
  const getSketchDataUrl = (): string => {
    const canvas = canvasRef.current;
    if (!canvas) return '';
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    if (!ctx) return '';
    const { data } = ctx.getImageData(0, 0, canvas.width, canvas.height);
    let blank = true;
    for (let i = 3; i < data.length; i += 4) {
      if (data[i] !== 0) { blank = false; break; }
    }
    return blank ? '' : canvas.toDataURL('image/png');
  };

  const getSavePayload = () => ({
    text,
    sketchDataUrl: getSketchDataUrl(),
    images,
    layerOrder,
  });

  return {
    text, setText, images, layerOrder, activeLayer, setActiveLayer,
    canvasRef, color, lineWidth, isErasing, eraserSize, eraserCursor,
    setColor, setLineWidth, setIsErasing, toggleErasing: () => setIsErasing((v) => !v), setEraserSize,
    clearCanvas, startDrawing, draw, stopDrawing,
    selectedImgId, setSelectedImgId,
    handleImageMouseDown, handleResizeMouseDown, deleteImage,
    handleUpload, handleFileChange, fileInputRef,
    moveLayer, reorderLayers, zIndex,
    getSketchDataUrl, getSavePayload,
    initCanvasForElement,
  };
}
