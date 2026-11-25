/** Shared workbook model. Rooms other than sheets will reference ranges, not copy this shape. */

export const ROW_COUNT = 100_000;
export const COL_COUNT = 256;
export const DEFAULT_ROW_HEIGHT = 26;
export const DEFAULT_COL_WIDTH = 112;
export const ROW_HEADER_WIDTH = 48;
export const COL_HEADER_HEIGHT = 28;

export interface CellRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface BorderEdge {
  style: 'thin' | 'medium' | 'thick';
  color: string;
}

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontFamily?: string;
  fontSize?: number;
  color?: string;
  fill?: string;
  align?: 'left' | 'center' | 'right';
  valign?: 'top' | 'middle' | 'bottom';
  wrap?: boolean;
  /** general, number, currency, percent, date, or a display hint. */
  format?: string;
  border?: Partial<Record<'top' | 'right' | 'bottom' | 'left', BorderEdge>>;
}

export interface CellRecord {
  input?: string;
  style?: CellStyle;
}

export type CondType = 'gt' | 'lt' | 'eq' | 'between' | 'contains' | 'colorScale';

export interface ConditionalRule {
  id: string;
  range: CellRange;
  type: CondType;
  value?: string;
  value2?: string;
  fill?: string;
  color?: string;
  minColor?: string;
  maxColor?: string;
}

export interface ValidationRule {
  id: string;
  range: CellRange;
  type: 'list';
  values: string[];
}

export type ChartType = 'bar' | 'line' | 'pie' | 'scatter';

export interface ChartModel {
  id: string;
  type: ChartType;
  range: CellRange;
  title: string;
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface FilterModel {
  range: CellRange;
  headerRow: number;
  /** Column -> display values that stay visible. An empty list hides every row. */
  criteria: Record<number, string[]>;
}

export interface NamedRange {
  name: string;
  sheetId: string;
  range: CellRange;
}

export interface SheetModel {
  id: string;
  name: string;
  cells: Record<string, CellRecord>;
  rowHeights: Record<number, number>;
  colWidths: Record<number, number>;
  hiddenRows: number[];
  hiddenCols: number[];
  filterHidden: number[];
  merges: CellRange[];
  freeze: { rows: number; cols: number };
  conditional: ConditionalRule[];
  validations: ValidationRule[];
  charts: ChartModel[];
  filter: FilterModel | null;
}

export interface WorkbookModel {
  id: string;
  title: string;
  createdAt: number;
  updatedAt: number;
  activeSheetId: string;
  sheets: SheetModel[];
  names: NamedRange[];
  x: number;
  y: number;
}

export type PlainValue = string | number | boolean | null;

export function keyOf(r: number, c: number): string {
  return `${r},${c}`;
}

export function parseKey(key: string): { r: number; c: number } {
  const i = key.indexOf(',');
  return { r: Number(key.slice(0, i)), c: Number(key.slice(i + 1)) };
}

export function normalizeRange(range: CellRange): CellRange {
  return {
    r1: Math.min(range.r1, range.r2),
    c1: Math.min(range.c1, range.c2),
    r2: Math.max(range.r1, range.r2),
    c2: Math.max(range.c1, range.c2),
  };
}

export function rangesEqual(a: CellRange, b: CellRange): boolean {
  const x = normalizeRange(a);
  const y = normalizeRange(b);
  return x.r1 === y.r1 && x.c1 === y.c1 && x.r2 === y.r2 && x.c2 === y.c2;
}

export function rangeContains(range: CellRange, r: number, c: number): boolean {
  const n = normalizeRange(range);
  return r >= n.r1 && r <= n.r2 && c >= n.c1 && c <= n.c2;
}

export function clampIndex(n: number, max: number): number {
  if (n < 0) return 0;
  if (n >= max) return max - 1;
  return n;
}

export function cloneData<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

export function createSheet(name: string, id: string = crypto.randomUUID()): SheetModel {
  return {
    id,
    name,
    cells: {},
    rowHeights: {},
    colWidths: {},
    hiddenRows: [],
    hiddenCols: [],
    filterHidden: [],
    merges: [],
    freeze: { rows: 0, cols: 0 },
    conditional: [],
    validations: [],
    charts: [],
    filter: null,
  };
}

export function createWorkbook(title: string, x = 80, y = 72): WorkbookModel {
  const sheet = createSheet('Sheet1');
  const now = Date.now();
  return {
    id: crypto.randomUUID(),
    title,
    createdAt: now,
    updatedAt: now,
    activeSheetId: sheet.id,
    sheets: [sheet],
    names: [],
    x,
    y,
  };
}

export function cellAt(sheet: SheetModel, r: number, c: number): CellRecord | undefined {
  return sheet.cells[keyOf(r, c)];
}

export function findMerge(sheet: SheetModel, r: number, c: number): CellRange | null {
  for (const merge of sheet.merges) {
    if (rangeContains(merge, r, c)) return merge;
  }
  return null;
}

export function usedBounds(sheet: SheetModel): CellRange {
  let r2 = 0;
  let c2 = 0;
  let any = false;
  for (const key of Object.keys(sheet.cells)) {
    const cell = sheet.cells[key];
    if (!cell?.input && !cell?.style) continue;
    const pos = parseKey(key);
    any = true;
    if (pos.r > r2) r2 = pos.r;
    if (pos.c > c2) c2 = pos.c;
  }
  for (const merge of sheet.merges) {
    any = true;
    if (merge.r2 > r2) r2 = merge.r2;
    if (merge.c2 > c2) c2 = merge.c2;
  }
  for (const chart of sheet.charts) {
    any = true;
    if (chart.range.r2 > r2) r2 = chart.range.r2;
    if (chart.range.c2 > c2) c2 = chart.range.c2;
  }
  if (!any) return { r1: 0, c1: 0, r2: 0, c2: 0 };
  return { r1: 0, c1: 0, r2, c2 };
}

export function uniqueSheetName(sheets: SheetModel[], base: string): string {
  const names = new Set(sheets.map((sheet) => sheet.name.toLowerCase()));
  if (!names.has(base.toLowerCase())) return base;
  let n = 2;
  while (names.has(`${base} ${n}`.toLowerCase())) n += 1;
  return `${base} ${n}`;
}

export function sanitizeSheetName(name: string): string {
  const cleaned = name.replace(/[\\/?*[\]:]/g, '').trim().slice(0, 31);
  return cleaned || 'Sheet';
}
