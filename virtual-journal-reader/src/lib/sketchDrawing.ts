import type { MouseEvent, TouchEvent } from 'react';

export interface Point {
  x: number;
  y: number;
}

export interface CanvasHistory {
  undoStack: ImageData[];
  redoStack: ImageData[];
  undoBytes: number;
  redoBytes: number;
}

const GUIDE_SNAP_STEP_DEG = 45;
const GUIDE_SNAP_THRESHOLD_DEG = 4;
const MAX_HISTORY_BYTES = 160 * 1024 * 1024;
const MAX_HISTORY_STATES = 160;

export function createCanvasHistory(): CanvasHistory {
  return {
    undoStack: [],
    redoStack: [],
    undoBytes: 0,
    redoBytes: 0,
  };
}

function snapshotBytes(snapshot: ImageData): number {
  return snapshot.data.byteLength;
}

function trimHistory(history: CanvasHistory): void {
  while (
    history.undoStack.length + history.redoStack.length > MAX_HISTORY_STATES ||
    history.undoBytes + history.redoBytes > MAX_HISTORY_BYTES
  ) {
    if (history.undoStack.length > 0) {
      const removed = history.undoStack.shift();
      if (removed) history.undoBytes -= snapshotBytes(removed);
    } else {
      const removed = history.redoStack.shift();
      if (removed) history.redoBytes -= snapshotBytes(removed);
    }
  }
}

export function pushUndoSnapshot(history: CanvasHistory, snapshot: ImageData | null): void {
  if (!snapshot) return;
  history.undoStack.push(snapshot);
  history.undoBytes += snapshotBytes(snapshot);
  history.redoStack = [];
  history.redoBytes = 0;
  trimHistory(history);
}

function pushRedoSnapshot(history: CanvasHistory, snapshot: ImageData | null): void {
  if (!snapshot) return;
  history.redoStack.push(snapshot);
  history.redoBytes += snapshotBytes(snapshot);
  trimHistory(history);
}

function pushUndoSnapshotPreserveRedo(history: CanvasHistory, snapshot: ImageData | null): void {
  if (!snapshot) return;
  history.undoStack.push(snapshot);
  history.undoBytes += snapshotBytes(snapshot);
  trimHistory(history);
}

function popUndoSnapshot(history: CanvasHistory): ImageData | null {
  const snapshot = history.undoStack.pop();
  if (!snapshot) return null;
  history.undoBytes -= snapshotBytes(snapshot);
  return snapshot;
}

function popRedoSnapshot(history: CanvasHistory): ImageData | null {
  const snapshot = history.redoStack.pop();
  if (!snapshot) return null;
  history.redoBytes -= snapshotBytes(snapshot);
  return snapshot;
}

function normalizeDegrees(deg: number): number {
  return ((deg % 360) + 360) % 360;
}

function angularDistanceDegrees(a: number, b: number): number {
  const diff = Math.abs(normalizeDegrees(a - b));
  return diff > 180 ? 360 - diff : diff;
}

export function straightLineTargetWithGuideSnap(
  origin: Point,
  target: Point,
  snapThresholdDeg = GUIDE_SNAP_THRESHOLD_DEG,
): Point {
  const dx = target.x - origin.x;
  const dy = target.y - origin.y;
  const dist = Math.hypot(dx, dy);
  if (dist === 0) return { x: origin.x, y: origin.y };

  const angleDeg = normalizeDegrees(Math.atan2(dy, dx) * 180 / Math.PI);
  const guideDeg = Math.round(angleDeg / GUIDE_SNAP_STEP_DEG) * GUIDE_SNAP_STEP_DEG;
  if (angularDistanceDegrees(angleDeg, guideDeg) > snapThresholdDeg) {
    return target;
  }

  const guideRad = (guideDeg * Math.PI) / 180;
  return {
    x: origin.x + Math.cos(guideRad) * dist,
    y: origin.y + Math.sin(guideRad) * dist,
  };
}

export function captureCanvasSnapshot(
  ctx: CanvasRenderingContext2D,
  canvas: HTMLCanvasElement,
): ImageData | null {
  if (canvas.width === 0 || canvas.height === 0) return null;
  try {
    return ctx.getImageData(0, 0, canvas.width, canvas.height);
  } catch {
    return null;
  }
}

export function restoreCanvasSnapshot(
  ctx: CanvasRenderingContext2D,
  snapshot: ImageData,
  canvas?: HTMLCanvasElement,
): void {
  if (!canvas || (snapshot.width === canvas.width && snapshot.height === canvas.height)) {
    ctx.putImageData(snapshot, 0, 0);
    return;
  }

  const sourceCanvas = document.createElement('canvas');
  sourceCanvas.width = snapshot.width;
  sourceCanvas.height = snapshot.height;
  const sourceCtx = sourceCanvas.getContext('2d');
  if (!sourceCtx) return;
  sourceCtx.putImageData(snapshot, 0, 0);

  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(sourceCanvas, 0, 0, canvas.width, canvas.height);
}

export function undoCanvas(
  ctx: CanvasRenderingContext2D,
  canvas: HTMLCanvasElement,
  history: CanvasHistory,
): boolean {
  const previous = popUndoSnapshot(history);
  if (!previous) return false;
  pushRedoSnapshot(history, captureCanvasSnapshot(ctx, canvas));
  restoreCanvasSnapshot(ctx, previous, canvas);
  return true;
}

export function redoCanvas(
  ctx: CanvasRenderingContext2D,
  canvas: HTMLCanvasElement,
  history: CanvasHistory,
): boolean {
  const next = popRedoSnapshot(history);
  if (!next) return false;
  pushUndoSnapshotPreserveRedo(history, captureCanvasSnapshot(ctx, canvas));
  restoreCanvasSnapshot(ctx, next, canvas);
  return true;
}

export function isShiftStraightLine(e: MouseEvent | TouchEvent): boolean {
  return !('touches' in e) && (e as MouseEvent).shiftKey;
}
