import { COL_COUNT, ROW_COUNT } from './types';
import { parseRef, rewriteFormulaRefs, writeRef, type ParsedRef } from './a1';

/** HyperFormula treats TRUE and FALSE as names unless they are called as functions. */
export function normalizeFormula(formula: string): string {
  if (!formula.startsWith('=')) return formula;
  let out = '';
  let i = 0;
  while (i < formula.length) {
    if (formula[i] === '"') {
      let j = i + 1;
      while (j < formula.length) {
        if (formula[j] === '"' && formula[j + 1] === '"') {
          j += 2;
          continue;
        }
        if (formula[j] === '"') {
          j += 1;
          break;
        }
        j += 1;
      }
      out += formula.slice(i, j);
      i = j;
      continue;
    }
    if (/[A-Za-z]/.test(formula[i] ?? '')) {
      let j = i + 1;
      while (j < formula.length && /[A-Za-z0-9_.]/.test(formula[j] ?? '')) j += 1;
      const word = formula.slice(i, j);
      const upper = word.toUpperCase();
      if ((upper === 'TRUE' || upper === 'FALSE') && formula[j] !== '(') {
        out += `${upper}()`;
      } else {
        out += word;
      }
      i = j;
      continue;
    }
    out += formula[i];
    i += 1;
  }
  return out;
}

function shiftRef(ref: ParsedRef, dRow: number, dCol: number, respectLocks: boolean): string | null {
  let row = ref.row;
  let col = ref.col;
  if (!respectLocks || !ref.rowAbs) row += dRow;
  if (!respectLocks || !ref.colAbs) col += dCol;
  if (row < 0 || col < 0 || row >= ROW_COUNT || col >= COL_COUNT) return null;
  return writeRef({ ...ref, row, col });
}

/** Copy, paste, and fill. Dollar locks stay put. */
export function offsetFormula(formula: string, dRow: number, dCol: number): string {
  if (!formula.startsWith('=') || (dRow === 0 && dCol === 0)) return formula;
  return rewriteFormulaRefs(formula, (token) => {
    const bits = token.split(':');
    const next = bits.map((bit) => {
      const ref = parseRef(bit);
      if (!ref) return bit;
      return shiftRef(ref, dRow, dCol, true) ?? '#REF!';
    });
    if (next.some((bit) => bit === '#REF!')) return '#REF!';
    return next.join(':');
  });
}

function mapStructural(index: number, at: number, count: number): number | null {
  if (count >= 0) return index >= at ? index + count : index;
  const removed = -count;
  if (index >= at && index < at + removed) return null;
  if (index >= at + removed) return index + count;
  return index;
}

/**
 * Row or column insert/delete. Absolute refs move too, because the cells moved.
 * `count` is positive for insert and negative for delete.
 */
export function adjustFormulaForStructure(
  formula: string,
  homeSheet: string,
  targetSheet: string,
  axis: 'row' | 'col',
  at: number,
  count: number,
): string {
  if (!formula.startsWith('=') || count === 0) return formula;
  return rewriteFormulaRefs(formula, (token) => {
    const bits = token.split(':');
    const next = bits.map((bit) => {
      const ref = parseRef(bit);
      if (!ref) return bit;
      const sheet = ref.sheet ?? homeSheet;
      if (sheet !== targetSheet) return bit;
      const current = axis === 'row' ? ref.row : ref.col;
      const mapped = mapStructural(current, at, count);
      if (mapped == null) return '#REF!';
      const updated = axis === 'row' ? { ...ref, row: mapped } : { ...ref, col: mapped };
      return writeRef(updated);
    });
    if (next.some((bit) => bit === '#REF!')) return '#REF!';
    return next.join(':');
  });
}

export function mapIndex(index: number, at: number, count: number): number | null {
  return mapStructural(index, at, count);
}
