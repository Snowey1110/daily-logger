import React, { useRef, useEffect, useState, useCallback } from 'react';
import { X, Eraser } from 'lucide-react';
import { useReaderT } from '../readerI18n';

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

export const DrawingCanvas: React.FC<DrawingCanvasProps> = ({ onSave, onClose, initialData, sketchId }) => {
  const { t } = useReaderT();
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const contextRef = useRef<CanvasRenderingContext2D | null>(null);
  const [isDrawing, setIsDrawing] = useState(false);
  const [color, setColor] = useState('#000000');
  const [lineWidth, setLineWidth] = useState(3);

  const initCanvas = useCallback(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;

    const { width, height } = canvas.getBoundingClientRect();
    if (width === 0 || height === 0) return;
    
    // Use a temporary canvas to preserve content during resize
    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = canvas.width;
    tempCanvas.height = canvas.height;
    const tempCtx = tempCanvas.getContext('2d');
    if (tempCtx && canvas.width > 0 && canvas.height > 0) {
      tempCtx.drawImage(canvas, 0, 0);
    }

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
    initCanvas();
    // Initial content load if any
    if (initialData && canvasRef.current) {
        const ctx = canvasRef.current.getContext('2d');
        if (ctx) {
            const img = new Image();
            img.onload = () => ctx.drawImage(img, 0, 0);
            img.src = initialData;
        }
    }

    window.addEventListener('resize', initCanvas);
    return () => window.removeEventListener('resize', initCanvas);
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
    
    let clientX, clientY;
    if ('touches' in e) {
      clientX = e.touches[0].clientX;
      clientY = e.touches[0].clientY;
    } else {
      clientX = e.clientX;
      clientY = e.clientY;
    }
    
    return {
      x: clientX - rect.left,
      y: clientY - rect.top
    };
  };

  const startDrawing = (e: React.MouseEvent | React.TouchEvent) => {
    if ('touches' in e) e.preventDefault();
    const { x, y } = getCoordinates(e);
    contextRef.current?.beginPath();
    contextRef.current?.moveTo(x, y);
    setIsDrawing(true);
  };

  const draw = (e: React.MouseEvent | React.TouchEvent) => {
    if (!isDrawing) return;
    if ('touches' in e) e.preventDefault();
    const { x, y } = getCoordinates(e);
    contextRef.current?.lineTo(x, y);
    contextRef.current?.stroke();
  };

  const stopDrawing = () => {
    if (isDrawing) {
      contextRef.current?.closePath();
      setIsDrawing(false);
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
      contextRef.current.clearRect(0, 0, canvas.width, canvas.height);
    }
  };

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
        
        <div className="flex-1 bg-[#2c1e14]/5 relative overflow-hidden flex items-center justify-center drawing-container p-8">
           <canvas
            ref={canvasRef}
            onMouseDown={startDrawing}
            onMouseMove={draw}
            onMouseUp={stopDrawing}
            onMouseOut={stopDrawing}
            onTouchStart={startDrawing}
            onTouchMove={draw}
            onTouchEnd={stopDrawing}
            className="bg-white shadow-[0_20px_50px_rgba(0,0,0,0.1)] cursor-crosshair touch-none rounded-lg"
            style={{ width: '100%', height: '100%' }}
          />
        </div>
      </div>
    </div>
  );
};
