/** Pixel box for one journal text column (inner body, after headers). */
export type JournalColumnMeasure = {
  widthPx: number;
  firstPageInnerHeightPx: number;
  restPageInnerHeightPx: number;
};

/** Synced from layout before spreads recompute (refs valid during split). */
export const journalMeasureRuntime = {
  styleSourceEl: null as HTMLElement | null,
  dims: null as JournalColumnMeasure | null,
};

const FALLBACK_MAX_CHARS = 900;

function splitTextFallback(text: string): string[] {
  if (!text) return [''];
  const pages: string[] = [];
  let currentText = text;
  while (currentText.length > 0) {
    if (currentText.length <= FALLBACK_MAX_CHARS) {
      pages.push(currentText);
      break;
    }
    let sliceIndex = currentText.lastIndexOf(' ', FALLBACK_MAX_CHARS);
    if (sliceIndex === -1) sliceIndex = FALLBACK_MAX_CHARS;
    pages.push(currentText.substring(0, sliceIndex).trim());
    currentText = currentText.substring(sliceIndex).trim();
  }
  return pages;
}

function snapToWordBoundary(text: string, cut: number, minProgress: number): number {
  if (cut <= 0) return Math.min(text.length, 1);
  if (cut >= text.length) return cut;
  const from = Math.max(0, cut - 500);
  const slice = text.slice(from, cut);
  const nl = slice.lastIndexOf('\n');
  if (nl >= 0 && from + nl + 1 >= minProgress) return from + nl + 1;
  const sp = slice.lastIndexOf(' ');
  if (sp >= 0 && from + sp + 1 >= minProgress) return from + sp + 1;
  return cut;
}

function applyTextStylesFromSource(target: HTMLElement, source: HTMLElement): void {
  const cs = getComputedStyle(source);
  const keys = [
    'font',
    'font-size',
    'font-family',
    'font-weight',
    'font-style',
    'line-height',
    'letter-spacing',
    'word-spacing',
    'font-variant',
    'text-rendering',
  ] as const;
  for (const k of keys) {
    target.style.setProperty(k, cs.getPropertyValue(k));
  }
  target.style.whiteSpace = 'pre-wrap';
  target.style.wordBreak = 'break-word';
  target.style.boxSizing = 'border-box';
}

/** Match Page.tsx: paragraphs with ~space-y-3 between blocks */
function fillParagraphs(container: HTMLElement, raw: string): void {
  container.textContent = '';
  const parts = raw.split('\n');
  const gap = 12;
  parts.forEach((para, i) => {
    const p = document.createElement('p');
    p.style.margin = '0';
    p.style.marginBottom = i < parts.length - 1 ? `${gap}px` : '0';
    p.textContent = para;
    container.appendChild(p);
  });
}

function scrollHeightFits(container: HTMLElement, maxH: number): boolean {
  return container.scrollHeight <= maxH + 2;
}

/**
 * Longest prefix of `text` that fits in a column of widthPx × maxHeightPx
 * using typography copied from the live journal text body.
 */
function takeFittingPrefixEnd(
  text: string,
  widthPx: number,
  maxHeightPx: number,
  styleSource: HTMLElement,
): number {
  if (!text.length) return 0;
  if (maxHeightPx < 24 || widthPx < 64) return Math.min(FALLBACK_MAX_CHARS, text.length);

  const probe = document.createElement('div');
  applyTextStylesFromSource(probe, styleSource);
  probe.style.position = 'fixed';
  probe.style.left = '-12000px';
  probe.style.top = '0';
  probe.style.width = `${widthPx}px`;
  probe.style.maxHeight = `${maxHeightPx}px`;
  probe.style.overflow = 'hidden';
  probe.style.visibility = 'hidden';
  probe.style.pointerEvents = 'none';
  document.body.appendChild(probe);

  const fitsLength = (n: number) => {
    const slice = text.slice(0, n);
    fillParagraphs(probe, slice);
    return scrollHeightFits(probe, maxHeightPx);
  };

  let end: number;
  if (fitsLength(text.length)) {
    end = text.length;
  } else {
    let lo = 0;
    let hi = text.length;
    while (lo < hi) {
      const mid = Math.floor((lo + hi + 1) / 2);
      if (fitsLength(mid)) lo = mid;
      else hi = mid - 1;
    }
    let cut = Math.max(1, lo);
    cut = snapToWordBoundary(text, cut, 1);
    if (!fitsLength(cut)) {
      cut = lo;
    }
    end = cut;
  }

  document.body.removeChild(probe);
  return end;
}

/** Split full journal into page strings using measured column heights (first vs continuation). */
export function splitTextIntoJournalPages(text: string): string[] {
  if (!text) return [''];
  const trimmed = text.trim();
  if (!trimmed.length) return [''];

  const { dims, styleSourceEl } = journalMeasureRuntime;
  if (!dims || !styleSourceEl || typeof document === 'undefined') {
    return splitTextFallback(text);
  }

  const { widthPx, firstPageInnerHeightPx, restPageInnerHeightPx } = dims;
  if (widthPx < 32 || firstPageInnerHeightPx < 24 || restPageInnerHeightPx < 24) {
    return splitTextFallback(text);
  }

  const pages: string[] = [];
  let rest = text;
  let pageIdx = 0;

  while (rest.length > 0) {
    const maxH = pageIdx === 0 ? firstPageInnerHeightPx : restPageInnerHeightPx;
    const end = takeFittingPrefixEnd(rest, widthPx, maxH, styleSourceEl);
    if (end <= 0) {
      pages.push(rest.charAt(0));
      rest = rest.slice(1);
      pageIdx++;
      continue;
    }
    const chunk = rest.slice(0, end).trimEnd();
    if (!chunk.length) {
      rest = rest.slice(end).trimStart();
      continue;
    }
    pages.push(chunk);
    rest = rest.slice(end).trimStart();
    pageIdx++;
    if (pageIdx > 2000) break;
  }

  return pages.length ? pages : [''];
}

/** Character-based pagination for STT/AI columns (non-journal). */
export function splitTextIntoPagesLegacy(text: string): string[] {
  return splitTextFallback(text);
}
