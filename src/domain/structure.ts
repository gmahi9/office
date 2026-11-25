import type { CellRange, CellRecord, SheetModel } from './types';
import { keyOf, parseKey } from './types';
import { mapIndex } from './formula';

export function shiftRange(range: CellRange, axis: 'row' | 'col', at: number, count: number): CellRange | null {
  const start = axis === 'row' ? range.r1 : range.c1;
  const end = axis === 'row' ? range.r2 : range.c2;
  if (count >= 0) {
    const ns = start >= at ? start + count : start;
    const ne = end >= at ? end + count : end;
    return applyEnds(range, axis, ns, ne);
  }
  const removed = -count;
  const endDel = at + removed - 1;
  if (end < at || start > endDel) {
    const ns = mapIndex(start, at, count);
    const ne = mapIndex(end, at, count);
    if (ns == null || ne == null) return null;
    return applyEnds(range, axis, ns, ne);
  }
  if (start >= at && end <= endDel) return null;
  if (start < at) {
    const ne = end > endDel ? end + count : at - 1;
    if (ne < start) return null;
    return applyEnds(range, axis, start, ne);
  }
  const ne = end + count;
  if (ne < at) return null;
  return applyEnds(range, axis, at, ne);
}

function applyEnds(range: CellRange, axis: 'row' | 'col', start: number, end: number): CellRange {
  if (axis === 'row') return { ...range, r1: start, r2: end };
  return { ...range, c1: start, c2: end };
}

export function shiftCells(
  cells: Record<string, CellRecord>,
  axis: 'row' | 'col',
  at: number,
  count: number,
): Record<string, CellRecord> {
  const next: Record<string, CellRecord> = {};
  for (const [key, cell] of Object.entries(cells)) {
    const pos = parseKey(key);
    const mapped = mapIndex(axis === 'row' ? pos.r : pos.c, at, count);
    if (mapped == null) continue;
    const r = axis === 'row' ? mapped : pos.r;
    const c = axis === 'col' ? mapped : pos.c;
    next[keyOf(r, c)] = cell;
  }
  return next;
}

export function shiftNumberMap(
  map: Record<number, number>,
  at: number,
  count: number,
): Record<number, number> {
  const next: Record<number, number> = {};
  for (const [key, value] of Object.entries(map)) {
    const mapped = mapIndex(Number(key), at, count);
    if (mapped == null) continue;
    next[mapped] = value;
  }
  return next;
}

export function shiftIndexList(list: number[], at: number, count: number): number[] {
  const out: number[] = [];
  for (const value of list) {
    const mapped = mapIndex(value, at, count);
    if (mapped != null) out.push(mapped);
  }
  return out;
}

export function shiftSheetGeometry(sheet: SheetModel, axis: 'row' | 'col', at: number, count: number): void {
  sheet.cells = shiftCells(sheet.cells, axis, at, count);
  if (axis === 'row') {
    sheet.rowHeights = shiftNumberMap(sheet.rowHeights, at, count);
    sheet.hiddenRows = shiftIndexList(sheet.hiddenRows, at, count);
    sheet.filterHidden = shiftIndexList(sheet.filterHidden, at, count);
    if (sheet.freeze.rows > at) {
      sheet.freeze = { ...sheet.freeze, rows: Math.max(at, sheet.freeze.rows + count) };
    }
  } else {
    sheet.colWidths = shiftNumberMap(sheet.colWidths, at, count);
    sheet.hiddenCols = shiftIndexList(sheet.hiddenCols, at, count);
    if (sheet.freeze.cols > at) {
      sheet.freeze = { ...sheet.freeze, cols: Math.max(at, sheet.freeze.cols + count) };
    }
  }
  sheet.merges = sheet.merges
    .map((merge) => shiftRange(merge, axis, at, count))
    .filter((merge): merge is CellRange => merge != null);
  sheet.conditional = sheet.conditional.flatMap((rule) => {
    const range = shiftRange(rule.range, axis, at, count);
    return range ? [{ ...rule, range }] : [];
  });
  sheet.validations = sheet.validations.flatMap((rule) => {
    const range = shiftRange(rule.range, axis, at, count);
    return range ? [{ ...rule, range }] : [];
  });
  sheet.charts = sheet.charts.flatMap((chart) => {
    const range = shiftRange(chart.range, axis, at, count);
    return range ? [{ ...chart, range }] : [];
  });
  if (sheet.filter) {
    const range = shiftRange(sheet.filter.range, axis, at, count);
    if (!range) {
      sheet.filter = null;
      sheet.filterHidden = [];
    } else {
      const header = mapIndex(sheet.filter.headerRow, at, count);
      sheet.filter = {
        ...sheet.filter,
        range,
        headerRow: header ?? range.r1,
      };
    }
  }
}
