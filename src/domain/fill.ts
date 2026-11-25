import type { CellRange, CellRecord, SheetModel } from './types';
import { cellAt, keyOf, normalizeRange } from './types';
import { offsetFormula } from './formula';

export interface FillEdit {
  r: number;
  c: number;
  record: CellRecord | null;
}

function stepOf(values: number[]): number | null {
  if (values.length < 2) return null;
  const step = values[1] - values[0];
  for (let i = 2; i < values.length; i += 1) {
    if (Math.abs(values[i] - values[i - 1] - step) > 1e-9) return null;
  }
  return step;
}

function numericInput(input: string | undefined): number | null {
  if (!input || input.startsWith('=')) return null;
  if (!/^-?\d+(\.\d+)?$/.test(input)) return null;
  const n = Number(input);
  return Number.isFinite(n) ? n : null;
}

/** Extend `source` to `target`. Target includes the original selection. */
export function computeFill(sheet: SheetModel, source: CellRange, target: CellRange): FillEdit[] {
  const src = normalizeRange(source);
  const dst = normalizeRange(target);
  const down = dst.r2 - src.r2;
  const up = src.r1 - dst.r1;
  const left = src.c1 - dst.c1;
  let dr = 0;
  let dc = 0;
  if (down > 0 || up > 0 || dst.c2 - src.c2 > 0 || left > 0) {
    const vertical = Math.max(down, up);
    const horizontal = Math.max(dst.c2 - src.c2, left);
    if (vertical >= horizontal && vertical > 0) dr = down >= up ? 1 : -1;
    else if (horizontal > 0) dc = dst.c2 - src.c2 >= left ? 1 : -1;
  }
  if (dr === 0 && dc === 0) return [];

  const height = src.r2 - src.r1 + 1;
  const width = src.c2 - src.c1 + 1;
  const edits: FillEdit[] = [];
  const seriesCols = new Map<number, number>();
  const seriesRows = new Map<number, number>();
  if (dr !== 0 && width >= 1 && height >= 2) {
    for (let c = src.c1; c <= src.c2; c += 1) {
      const values: number[] = [];
      for (let r = src.r1; r <= src.r2; r += 1) values.push(numericInput(cellAt(sheet, r, c)?.input) ?? Number.NaN);
      if (values.every((n) => Number.isFinite(n))) {
        const step = stepOf(values);
        if (step != null) seriesCols.set(c, step);
      }
    }
  }
  if (dc !== 0 && height >= 1 && width >= 2) {
    for (let r = src.r1; r <= src.r2; r += 1) {
      const values: number[] = [];
      for (let c = src.c1; c <= src.c2; c += 1) values.push(numericInput(cellAt(sheet, r, c)?.input) ?? Number.NaN);
      if (values.every((n) => Number.isFinite(n))) {
        const step = stepOf(values);
        if (step != null) seriesRows.set(r, step);
      }
    }
  }

  const rStart = dr > 0 ? src.r2 + 1 : dst.r1;
  const rEnd = dr > 0 ? dst.r2 : src.r1 - 1;
  const cStart = dc > 0 ? src.c2 + 1 : dst.c1;
  const cEnd = dc > 0 ? dst.c2 : src.c1 - 1;
  const rowFrom = dr !== 0 ? Math.min(rStart, rEnd) : src.r1;
  const rowTo = dr !== 0 ? Math.max(rStart, rEnd) : src.r2;
  const colFrom = dc !== 0 ? Math.min(cStart, cEnd) : src.c1;
  const colTo = dc !== 0 ? Math.max(cStart, cEnd) : src.c2;

  for (let r = rowFrom; r <= rowTo; r += 1) {
    for (let c = colFrom; c <= colTo; c += 1) {
      if (r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2) continue;
      const srcR = src.r1 + mod(r - src.r1, height);
      const srcC = src.c1 + mod(c - src.c1, width);
      const origin = cellAt(sheet, srcR, srcC);
      if (!origin || (!origin.input && !origin.style)) {
        edits.push({ r, c, record: null });
        continue;
      }
      let input = origin.input;
      if (dr !== 0 && seriesCols.has(c) && input && !input.startsWith('=')) {
        const step = seriesCols.get(c) ?? 0;
        const base = numericInput(cellAt(sheet, dr > 0 ? src.r2 : src.r1, c)?.input) ?? 0;
        const steps = dr > 0 ? r - src.r2 : src.r1 - r;
        input = String(base + steps * step * (dr > 0 ? 1 : -1));
      } else if (dc !== 0 && seriesRows.has(r) && input && !input.startsWith('=')) {
        const step = seriesRows.get(r) ?? 0;
        const base = numericInput(cellAt(sheet, r, dc > 0 ? src.c2 : src.c1)?.input) ?? 0;
        const steps = dc > 0 ? c - src.c2 : src.c1 - c;
        input = String(base + steps * step * (dc > 0 ? 1 : -1));
      } else if (input?.startsWith('=')) {
        input = offsetFormula(input, r - srcR, c - srcC);
      }
      const record: CellRecord = {};
      if (input) record.input = input;
      if (origin.style) record.style = origin.style;
      edits.push({ r, c, record: record.input || record.style ? record : null });
    }
  }
  return edits;
}

function mod(n: number, m: number): number {
  return ((n % m) + m) % m;
}

export function fillDownRange(anchor: CellRange): CellRange {
  const src = normalizeRange(anchor);
  return { ...src, r2: src.r2 + (src.r2 - src.r1 + 1) };
}

export function fillRightRange(anchor: CellRange): CellRange {
  const src = normalizeRange(anchor);
  return { ...src, c2: src.c2 + (src.c2 - src.c1 + 1) };
}

export function editKey(edit: FillEdit): string {
  return keyOf(edit.r, edit.c);
}
