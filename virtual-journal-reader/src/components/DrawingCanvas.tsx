import React, { useRef, useEffect, useState, useCallback } from 'react';
import { X, Eraser } from 'lucide-react';
import { useReaderT } from '../readerI18n';
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

interface DrawingCanvasProps {
  onSave: (dataUrl: string, sketchId?: string) => void;
  onClose: () => void;
  initialData?: string;
  /** When editing an existing sketch, its ID is passed here. */
  sketchId?: string;
}

function isCanvasUniformBlank(canvas: HTMLCanvasElement): boolean {
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  if (!ctx || canvas.width === 0 || canvas.height === 0) {
    return true;
  }
  const { data } = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const r0 = data[0];
  const g0 = data[1];
  const b0 = data[2];
  const a0 = data[3];
  for (let i = 4; i < data.length; i += 4) {
    if (data[i] !== r0 || data[i + 1] !== g0 || data[i + 2] !== b0 || data[i + 3] !== a0) {
      return false;
    }
  }
  return true;
}

function getDrawableHostSize(host: HTMLElement): { width: number; height: number } {
  const styles = window.getComputedStyle(host);
  const horizontalPadding = parseFloat(styles.paddingLeft) + parseFloat(styles.paddingRight);
  const verticalPadding = parseFloat(styles.paddingTop) + parseFloat(styles.paddingBottom);
  const rect = host.getBoundingClientRect();
  return {
    width: Math.max(0, Math.round(rect.width - horizontalPadding)),
    height: Math.max(0, Math.round(rect.height - verticalPadding)),
  };
}

export const DrawingCanvas: React.FC<DrawingCanvasProps> = ({ onSave, onClose, initialData, sketchId }) => {
  const { t } = useReaderT();
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const contextRef = useRef<CanvasRenderingContext2D | null>(null);
  const strokeStartRef = useRef<{ x: number; y: number } | null>(null);
  const strokeSnapshotRef = useRef<ImageData | null>(null);
  const historyRef = useRef(createCanvasHistory());
  const [isDrawing, setIsDrawing] = useState(false);
  const [color, setColor] = useState('#000000');
  const [lineWidth, setLineWidth] = useState(3);

  const initCanvas = useCallback(() => {
    const canvas = canvasRef.current;
    const host = canvas?.parentElement;
    if (!canvas || !host) return;

    const { width, height } = getDrawableHostSize(host);
    if (width === 0 || height === 0) return;

    // Use a temporary canvas to preserve content during resize
    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = canvas.width;
    tempCanvas.height = canvas.height;
    const tempCtx = tempCanvas.getContext('2d');
    if (tempCtx && canvas.width > 0 && canvas.height > 0) {
      tempCtx.drawImage(canvas, 0, 0);
    }

    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;

    canvas.width = width;
    canvas.height = height;

    const ctx = canvas.getContext('2d');
    if (ctx) {
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.strokeStyle = color;
      ctx.lineWidth = lineWidth;
      contextRef.current = ctx;

      // Restore content
      if (tempCanvas.width > 0 && tempCanvas.height > 0) {
        ctx.drawImage(tempCanvas, 0, 0);
      } else if (initialData) {
        const img = new Image();
        img.onload = () => ctx.drawImage(img, 0, 0);
        img.src = initialData;
      }
    }
  }, [color, lineWidth, initialData]);

  useEffect(() => {
    const canvas = canvasRef.current;
    const host = canvas?.parentElement;
    if (!canvas) return;

    const runInit = () => initCanvas();
    runInit();
    const raf = requestAnimationFrame(() => {
      runInit();
      requestAnimationFrame(runInit);
    });

    const ro = host ? new ResizeObserver(runInit) : null;
    ro?.observe(host);
    window.addEventListener('resize', runInit);
    return () => {
      cancelAnimationFrame(raf);
      ro?.disconnect();
      window.removeEventListener('resize', runInit);
    };
  }, [initCanvas, initialData]);

  useEffect(() => {
    if (contextRef.current) {
      contextRef.current.strokeStyle = color;
      contextRef.current.lineWidth = lineWidth;
    }
  }, [color, lineWidth]);

  const getCoordinates = (e: React.MouseEvent | React.TouchEvent) => {
    const canvas = canvasRef.current;
    if (!canvas) return { x: 0, y: 0 };
    const rect = canvas.getBoundingClientRect();
    if (rect.width === 0 || rect.height === 0) return { x: 0, y: 0 };

    let clientX: number, clientY: number;
    if ('touches' in e) {
      clientX = e.touches[0].clientX;
      clientY = e.touches[0].clientY;
    } else {
      clientX = e.clientX;
      clientY = e.clientY;
    }

    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    return {
      x: (clientX - rect.left) * scaleX,
      y: (clientY - rect.top) * scaleY,
    };
  };

  const startDrawing = (e: React.MouseEvent | React.TouchEvent) => {
    if ('touches' in e) e.preventDefault();
    initCanvas();
    const { x, y } = getCoordinates(e);
    strokeStartRef.current = { x, y };
    const ctx = contextRef.current;
    const canvas = canvasRef.current;
    if (ctx && canvas) {
      strokeSnapshotRef.current = captureCanvasSnapshot(ctx, canvas);
      pushUndoSnapshot(historyRef.current, strokeSnapshotRef.current);
    }
    ctx?.beginPath();
    ctx?.moveTo(x, y);
    setIsDrawing(true);
  };

  const draw = (e: React.MouseEvent | React.TouchEvent) => {
    if (!isDrawing) return;
    if ('touches' in e) e.preventDefault();
    const { x, y } = getCoordinates(e);
    const ctx = contextRef.current;
    if (!ctx) return;

    const start = strokeStartRef.current;
    if (isShiftStraightLine(e) && start && strokeSnapshotRef.current) {
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
    if (isDrawing) {
      contextRef.current?.closePath();
      setIsDrawing(false);
      strokeStartRef.current = null;
      strokeSnapshotRef.current = null;
    }
  };

  const handleSave = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    if (isCanvasUniformBlank(canvas)) {
      onSave('', sketchId);
      return;
    }
    onSave(canvas.toDataURL(), sketchId);
  };

  const clearCanvas = () => {
    const canvas = canvasRef.current;
    if (canvas && contextRef.current) {
      pushUndoSnapshot(historyRef.current, captureCanvasSnapshot(contextRef.current, canvas));
      contextRef.current.clearRect(0, 0, canvas.width, canvas.height);
    }
  };

  const undoSketch = useCallback(() => {
    const canvas = canvasRef.current;
    const ctx = contextRef.current;
    if (!canvas || !ctx) return;
    undoCanvas(ctx, canvas, historyRef.current);
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = color;
    ctx.lineWidth = lineWidth;
  }, [color, lineWidth]);

  const redoSketch = useCallback(() => {
    const canvas = canvasRef.current;
    const ctx = contextRef.current;
    if (!canvas || !ctx) return;
    redoCanvas(ctx, canvas, historyRef.current);
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = color;
    ctx.lineWidth = lineWidth;
  }, [color, lineWidth]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
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
  }, [undoSketch, redoSketch]);

  return (
    <div className="fixed inset-0 z-[100] bg-black/80 flex items-center justify-center p-4 backdrop-blur-md font-sans">
      <div className="bg-[#fdfaf2] rounded-2xl shadow-[0_50px_100px_-20px_rgba(0,0,0,0.5)] w-full max-w-4xl flex flex-col overflow-hidden h-[80vh] border border-[#d9c5b2]/20">
        <div className="p-4 border-b border-[#d9c5b2]/10 flex items-center justify-between bg-[#fbf8ef]">
          <div className="flex items-center gap-4">
            <h3 className="font-semibold text-slate-800 uppercase tracking-widest text-xs">{t('sketchpad')}</h3>
            <div className="flex items-center gap-2 border-l border-[#d9c5b2]/20 pl-4">
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
                className="w-24 accent-slate-600"
              />
            </div>
            <span className="text-[10px] text-slate-500 border-l border-[#d9c5b2]/20 pl-4">
              {t('sketchShiftStraightLine')}
            </span>
          </div>
          <div className="flex items-center gap-3">
            <button 
              onClick={clearCanvas}
              className="p-2 hover:bg-black/5 rounded-full text-slate-600 transition-colors"
              title={t('clearCanvas')}
            >
              <Eraser size={18} />
            </button>
            <button 
              onClick={handleSave}
              className="px-6 py-2 bg-slate-800 text-[#d9c5b2] rounded-full hover:bg-slate-900 transition-all font-semibold text-xs tracking-widest uppercase shadow-lg shadow-black/10"
            >
              {t('saveSketch')}
            </button>
            <button 
              onClick={onClose}
              className="p-2 hover:bg-black/5 rounded-full text-slate-400 transition-colors"
            >
              <X size={20} />
            </button>
          </div>
        </div>
        
        <div className="flex-1 bg-[#2c1e14]/5 relative overflow-hidden drawing-container p-8">
           <canvas
            ref={canvasRef}
            onMouseDown={startDrawing}
            onMouseMove={draw}
            onMouseUp={stopDrawing}
            onMouseOut={stopDrawing}
            onTouchStart={startDrawing}
            onTouchMove={draw}
            onTouchEnd={stopDrawing}
            className="bg-white shadow-[0_20px_50px_rgba(0,0,0,0.1)] cursor-crosshair touch-none rounded-lg block"
          />
        </div>
      </div>
    </div>
  );
};
