import { COL_COUNT, ROW_COUNT, type CellRange } from './types';

export function colToA(col: number): string {
  let n = col + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function colFromA(label: string): number {
  let n = 0;
  const text = label.toUpperCase();
  for (let i = 0; i < text.length; i += 1) {
    n = n * 26 + (text.charCodeAt(i) - 64);
  }
  return n - 1;
}

export function quoteSheet(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

export function formatA1(r: number, c: number, absCol = false, absRow = false): string {
  return `${absCol ? '$' : ''}${colToA(c)}${absRow ? '$' : ''}${r + 1}`;
}

export function formatRangeA1(range: CellRange, absolute = false): string {
  const a = formatA1(range.r1, range.c1, absolute, absolute);
  const b = formatA1(range.r2, range.c2, absolute, absolute);
  return a === b ? a : `${a}:${b}`;
}

export interface ParsedRef {
  sheet: string | null;
  col: number;
  row: number;
  colAbs: boolean;
  rowAbs: boolean;
}

export function parseRef(token: string): ParsedRef | null {
  let sheet: string | null = null;
  let rest = token;
  if (token.startsWith("'")) {
    let j = 1;
    let name = '';
    while (j < token.length) {
      if (token[j] === "'" && token[j + 1] === "'") {
        name += "'";
        j += 2;
        continue;
      }
      if (token[j] === "'") {
        j += 1;
        break;
      }
      name += token[j];
      j += 1;
    }
    if (token[j] !== '!') return null;
    sheet = name;
    rest = token.slice(j + 1);
  } else {
    const bang = token.lastIndexOf('!');
    if (bang !== -1) {
      sheet = token.slice(0, bang);
      rest = token.slice(bang + 1);
    }
  }
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(rest);
  if (!match) return null;
  const col = colFromA(match[2]);
  const row = Number(match[4]) - 1;
  if (col < 0 || col >= COL_COUNT || row < 0 || row >= ROW_COUNT) return null;
  return {
    sheet,
    col,
    row,
    colAbs: match[1] === '$',
    rowAbs: match[3] === '$',
  };
}

export function writeRef(ref: ParsedRef, homeSheet?: string): string {
  const body = formatA1(ref.row, ref.col, ref.colAbs, ref.rowAbs);
  if (!ref.sheet || (homeSheet && ref.sheet === homeSheet)) return body;
  return `${quoteSheet(ref.sheet)}!${body}`;
}

/** Read one A1 ref, or a range of two refs, starting at `index`. */
export function readRefToken(formula: string, index: number): { token: string; range: boolean } | null {
  const first = readSingleRef(formula, index);
  if (!first) return null;
  if (formula[index + first.length] === ':') {
    const second = readSingleRef(formula, index + first.length + 1);
    if (second) {
      return { token: `${first}:${second}`, range: true };
    }
  }
  return { token: first, range: false };
}

function readSingleRef(formula: string, index: number): string | null {
  let j = index;
  if (formula[j] === "'") {
    j += 1;
    while (j < formula.length) {
      if (formula[j] === "'" && formula[j + 1] === "'") {
        j += 2;
        continue;
      }
      if (formula[j] === "'") {
        j += 1;
        break;
      }
      j += 1;
    }
    if (formula[j] !== '!') return null;
    j += 1;
  } else if (/[A-Za-z_]/.test(formula[j] ?? '')) {
    let k = j;
    while (k < formula.length && /[A-Za-z0-9_.]/.test(formula[k] ?? '')) k += 1;
    if (formula[k] === '!') j = k + 1;
  }
  const body = j;
  if (formula[j] === '$') j += 1;
  const colStart = j;
  while (j < formula.length && j - colStart < 3 && /[A-Za-z]/.test(formula[j] ?? '')) j += 1;
  if (j === colStart) return null;
  if (formula[j] === '$') j += 1;
  const rowStart = j;
  while (j < formula.length && /[0-9]/.test(formula[j] ?? '')) j += 1;
  if (j === rowStart) return null;
  if (j < formula.length && /[A-Za-z0-9_]/.test(formula[j] ?? '')) return null;
  if (j === body) return null;
  return formula.slice(index, j);
}

export function rewriteFormulaRefs(formula: string, mapToken: (token: string) => string): string {
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
    const prev = i === 0 ? '' : formula[i - 1];
    if (i === 0 || !/[A-Za-z0-9_]/.test(prev ?? '')) {
      const found = readRefToken(formula, i);
      if (found) {
        out += mapToken(found.token);
        i += found.token.length;
        continue;
      }
    }
    out += formula[i];
    i += 1;
  }
  return out;
}

export function parseA1Range(text: string): CellRange | null {
  const parts = text.split(':');
  if (parts.length > 2) return null;
  const start = parseRef(parts[0] ?? '');
  const end = parseRef(parts[1] ?? parts[0] ?? '');
  if (!start || !end) return null;
  return {
    r1: Math.min(start.row, end.row),
    c1: Math.min(start.col, end.col),
    r2: Math.max(start.row, end.row),
    c2: Math.max(start.col, end.col),
  };
}
