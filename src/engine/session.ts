import { HyperFormula } from 'hyperformula';
import { formatRangeA1, quoteSheet } from '../domain/a1';
import { paintForCell, scaleBounds, type Paint } from '../domain/conditional';
import {
  coerceEngineValue,
  formatValue,
  interpretInput,
  plainFromUnknown,
} from '../domain/format';
import { normalizeFormula, offsetFormula } from '../domain/formula';
import { shiftRange, shiftSheetGeometry } from '../domain/structure';
import {
  COL_COUNT,
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  ROW_COUNT,
  cellAt,
  cloneData,
  createSheet,
  findMerge,
  keyOf,
  normalizeRange,
  parseKey,
  rangeContains,
  sanitizeSheetName,
  uniqueSheetName,
  usedBounds,
  type CellRange,
  type CellRecord,
  type CellStyle,
  type ChartModel,
  type ChartType,
  type ConditionalRule,
  type NamedRange,
  type PlainValue,
  type SheetModel,
  type ValidationRule,
  type WorkbookModel,
} from '../domain/types';
import type { LinkedRangeCell } from '../data/linkedRanges';

function createEngine(): HyperFormula {
  return HyperFormula.buildEmpty({
    licenseKey: 'gpl-v3',
    maxRows: ROW_COUNT,
    maxColumns: COL_COUNT,
  });
}

function compareValues(a: PlainValue, b: PlainValue): number {
  const emptyA = a == null || a === '';
  const emptyB = b == null || b === '';
  if (emptyA && emptyB) return 0;
  if (emptyA) return 1;
  if (emptyB) return -1;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'number') return -1;
  if (typeof b === 'number') return 1;
  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
}

export class WorkbookSession {
  model: WorkbookModel;
  hf: HyperFormula;
  version = 0;
  dirty = false;
  private past: WorkbookModel[] = [];
  private future: WorkbookModel[] = [];
  private gesture = false;

  constructor(model: WorkbookModel) {
    this.model = model;
    this.hf = createEngine();
    this.hydrate();
  }

  destroy(): void {
    this.hf.destroy();
  }

  get canUndo(): boolean {
    return this.past.length > 0;
  }

  get canRedo(): boolean {
    return this.future.length > 0;
  }

  sheet(id: string): SheetModel | undefined {
    return this.model.sheets.find((item) => item.id === id);
  }

  activeSheet(): SheetModel {
    return this.sheet(this.model.activeSheetId) ?? this.model.sheets[0];
  }

  setActive(sheetId: string): void {
    if (!this.sheet(sheetId) || this.model.activeSheetId === sheetId) return;
    this.model.activeSheetId = sheetId;
    this.touch();
  }

  inputAt(sheetId: string, r: number, c: number): string {
    return this.sheet(sheetId)?.cells[keyOf(r, c)]?.input ?? '';
  }

  raw(sheetId: string, r: number, c: number): PlainValue {
    const sheet = this.sheet(sheetId);
    if (!sheet) return null;
    return plainFromUnknown(this.hf.getCellValue({ sheet: this.hfId(sheet), row: r, col: c }));
  }

  display(sheetId: string, r: number, c: number): string {
    const sheet = this.sheet(sheetId);
    if (!sheet) return '';
    const cell = cellAt(sheet, r, c);
    let format = cell?.style?.format;
    if (!format || format === 'general') {
      const kind = String(this.hf.getCellValueDetailedType({ sheet: this.hfId(sheet), row: r, col: c }));
      if (kind === 'NUMBER_DATE' || kind === 'NUMBER_DATETIME' || kind === 'NUMBER_TIME') format = 'date';
      else if (kind === 'NUMBER_PERCENT') format = 'percent';
      else if (kind === 'NUMBER_CURRENCY') format = 'currency';
    }
    return formatValue(this.raw(sheetId, r, c), format);
  }

  paint(sheetId: string, r: number, c: number, bounds: Map<string, { min: number; max: number }>): Paint {
    const sheet = this.sheet(sheetId);
    if (!sheet) return {};
    return paintForCell(sheet, r, c, this.raw(sheetId, r, c), bounds);
  }

  scaleMap(sheetId: string): Map<string, { min: number; max: number }> {
    const sheet = this.sheet(sheetId);
    const map = new Map<string, { min: number; max: number }>();
    if (!sheet) return map;
    for (const rule of sheet.conditional) {
      if (rule.type !== 'colorScale') continue;
      map.set(rule.id, scaleBounds(sheet, rule, (r, c) => this.raw(sheetId, r, c)));
    }
    return map;
  }

  setCell(sheetId: string, r: number, c: number, rawInput: string): boolean {
    const sheet = this.requireSheet(sheetId);
    const next = interpretInput(rawInput).stored;
    const current = sheet.cells[keyOf(r, c)]?.input ?? '';
    if (next === current) return false;
    this.checkpoint();
    this.writeCell(sheet, r, c, rawInput);
    this.touch();
    return true;
  }

  applyRecords(sheetId: string, edits: Array<{ r: number; c: number; record: CellRecord | null }>): void {
    if (!edits.length) return;
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    const hfId = this.hfId(sheet);
    this.hf.batch(() => {
      for (const edit of edits) {
        const key = keyOf(edit.r, edit.c);
        if (!edit.record || (!edit.record.input && !edit.record.style)) {
          delete sheet.cells[key];
          this.hf.setCellContents({ sheet: hfId, row: edit.r, col: edit.c }, null);
          continue;
        }
        sheet.cells[key] = edit.record;
        this.pushEngine(hfId, edit.r, edit.c, edit.record.input ?? '');
      }
    });
    this.touch();
  }

  clear(sheetId: string, range: CellRange, mode: 'all' | 'values' | 'formats'): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    const hfId = this.hfId(sheet);
    const area = normalizeRange(range);
    this.hf.batch(() => {
      for (const key of Object.keys(sheet.cells)) {
        const pos = parseKey(key);
        if (!rangeContains(area, pos.r, pos.c)) continue;
        const cell = sheet.cells[key];
        if (!cell) continue;
        if (mode === 'formats') {
          delete cell.style;
          if (!cell.input) delete sheet.cells[key];
          continue;
        }
        this.hf.setCellContents({ sheet: hfId, row: pos.r, col: pos.c }, null);
        if (mode === 'all') delete sheet.cells[key];
        else {
          delete cell.input;
          if (!cell.style) delete sheet.cells[key];
        }
      }
    });
    this.touch();
  }

  patchStyle(sheetId: string, range: CellRange, patch: Partial<CellStyle>): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    const area = normalizeRange(range);
    this.eachCell(sheet, area, (r, c) => {
      const key = keyOf(r, c);
      const cell = sheet.cells[key] ?? {};
      const style: CellStyle = { ...(cell.style ?? {}), ...patch };
      if (patch.format === 'general') delete style.format;
      for (const prop of Object.keys(style) as (keyof CellStyle)[]) {
        if (style[prop] === undefined) delete style[prop];
      }
      if (Object.keys(style).length) cell.style = style;
      else delete cell.style;
      if (cell.input || cell.style) sheet.cells[key] = cell;
      else delete sheet.cells[key];
    });
    this.touch();
  }

  applyBorder(sheetId: string, range: CellRange, mode: 'all' | 'outer' | 'bottom' | 'none', color = '#1c2430'): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    const area = normalizeRange(range);
    const edge = { style: 'thin' as const, color };
    this.eachCell(sheet, area, (r, c) => {
        const key = keyOf(r, c);
        const cell = sheet.cells[key] ?? {};
        const border = { ...(cell.style?.border ?? {}) };
        if (mode === 'none') {
          if (cell.style) delete cell.style.border;
        } else {
          if (mode === 'all' || (mode === 'outer' && r === area.r1)) border.top = edge;
          if (mode === 'all' || (mode === 'outer' && r === area.r2) || (mode === 'bottom' && r === area.r2)) border.bottom = edge;
          if (mode === 'all' || (mode === 'outer' && c === area.c1)) border.left = edge;
          if (mode === 'all' || (mode === 'outer' && c === area.c2)) border.right = edge;
          cell.style = { ...(cell.style ?? {}), border };
        }
        if (cell.style && Object.keys(cell.style).length === 0) delete cell.style;
        if (cell.input || cell.style) sheet.cells[key] = cell;
        else delete sheet.cells[key];
    });
    this.touch();
  }

  merge(sheetId: string, range: CellRange): void {
    const area = normalizeRange(range);
    if (area.r1 === area.r2 && area.c1 === area.c2) {
      this.unmerge(sheetId, area);
      return;
    }
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    sheet.merges = sheet.merges.filter((merge) => !overlaps(merge, area));
    sheet.merges.push(area);
    this.touch();
  }

  unmerge(sheetId: string, range: CellRange): void {
    const sheet = this.requireSheet(sheetId);
    const area = normalizeRange(range);
    const next = sheet.merges.filter((merge) => !overlaps(merge, area));
    if (next.length === sheet.merges.length) return;
    this.checkpoint();
    sheet.merges = next;
    this.touch();
  }

  setFreeze(sheetId: string, rows: number, cols: number): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    sheet.freeze = {
      rows: Math.max(0, Math.min(rows, 20)),
      cols: Math.max(0, Math.min(cols, 10)),
    };
    this.touch();
  }

  setColWidth(sheetId: string, col: number, width: number): void {
    const sheet = this.requireSheet(sheetId);
    const next = Math.max(32, Math.min(480, Math.round(width)));
    if (Math.abs(next - DEFAULT_COL_WIDTH) < 2) delete sheet.colWidths[col];
    else sheet.colWidths[col] = next;
    this.mark();
  }

  setRowHeight(sheetId: string, row: number, height: number): void {
    const sheet = this.requireSheet(sheetId);
    const next = Math.max(18, Math.min(160, Math.round(height)));
    if (Math.abs(next - DEFAULT_ROW_HEIGHT) < 2) delete sheet.rowHeights[row];
    else sheet.rowHeights[row] = next;
    this.mark();
  }

  beginGesture(): void {
    if (this.gesture) return;
    this.checkpoint();
    this.gesture = true;
  }

  endGesture(): void {
    if (!this.gesture) return;
    this.gesture = false;
    this.touch();
  }

  hideRows(sheetId: string, rows: number[]): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    sheet.hiddenRows = [...new Set([...sheet.hiddenRows, ...rows])];
    this.touch();
  }

  hideCols(sheetId: string, cols: number[]): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    sheet.hiddenCols = [...new Set([...sheet.hiddenCols, ...cols])];
    this.touch();
  }

  unhideAll(sheetId: string): void {
    const sheet = this.requireSheet(sheetId);
    if (!sheet.hiddenRows.length && !sheet.hiddenCols.length) return;
    this.checkpoint();
    sheet.hiddenRows = [];
    sheet.hiddenCols = [];
    this.touch();
  }

  insertRows(sheetId: string, at: number, count = 1): boolean {
    return this.shiftAxis(sheetId, 'row', at, count);
  }

  deleteRows(sheetId: string, at: number, count = 1): boolean {
    return this.shiftAxis(sheetId, 'row', at, -count);
  }

  insertCols(sheetId: string, at: number, count = 1): boolean {
    return this.shiftAxis(sheetId, 'col', at, count);
  }

  deleteCols(sheetId: string, at: number, count = 1): boolean {
    return this.shiftAxis(sheetId, 'col', at, -count);
  }

  sort(sheetId: string, range: CellRange, column: number, direction: 'asc' | 'desc'): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    const area = normalizeRange(range);
    const header = sheet.filter && sheet.filter.range.r1 === area.r1 ? 1 : 0;
    const start = area.r1 + header;
    if (start > area.r2) return;
    const order: Array<{ r: number; value: PlainValue }> = [];
    for (let r = start; r <= area.r2; r += 1) order.push({ r, value: this.raw(sheetId, r, column) });
    const sorted = order
      .map((row, index) => ({ ...row, index }))
      .sort((a, b) => {
        const cmp = compareValues(a.value, b.value);
        if (cmp !== 0) return direction === 'asc' ? cmp : -cmp;
        return a.index - b.index;
      });
    const original = new Map<number, Map<number, CellRecord | undefined>>();
    for (const row of order) {
      const cols = new Map<number, CellRecord | undefined>();
      for (let c = area.c1; c <= area.c2; c += 1) cols.set(c, sheet.cells[keyOf(row.r, c)]);
      original.set(row.r, cols);
    }
    const hfId = this.hfId(sheet);
    this.hf.batch(() => {
      sorted.forEach((row, index) => {
        const dst = start + index;
        const delta = dst - row.r;
        const cols = original.get(row.r);
        for (let c = area.c1; c <= area.c2; c += 1) {
          const source = cols?.get(c);
          const key = keyOf(dst, c);
          if (!source || (!source.input && !source.style)) {
            delete sheet.cells[key];
            this.hf.setCellContents({ sheet: hfId, row: dst, col: c }, null);
            continue;
          }
          let input = source.input;
          if (input?.startsWith('=')) input = offsetFormula(input, delta, 0);
          const record: CellRecord = {};
          if (input) record.input = input;
          if (source.style) record.style = cloneData(source.style);
          sheet.cells[key] = record;
          this.pushEngine(hfId, dst, c, input ?? '');
        }
      });
    });
    this.recomputeFilter(sheet);
    this.touch();
  }

  toggleFilter(sheetId: string, range: CellRange): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    const area = normalizeRange(range);
    if (sheet.filter && sheet.filter.range.r1 === area.r1 && sheet.filter.range.c1 === area.c1 && sheet.filter.range.r2 === area.r2 && sheet.filter.range.c2 === area.c2) {
      sheet.filter = null;
      sheet.filterHidden = [];
    } else {
      const bounds = usedBounds(sheet);
      const next = area.r1 === area.r2 && area.c1 === area.c2
        ? { r1: bounds.r1, c1: bounds.c1, r2: Math.max(bounds.r2, area.r1), c2: Math.max(bounds.c2, area.c1) }
        : area;
      sheet.filter = { range: next, headerRow: next.r1, criteria: {} };
      sheet.filterHidden = [];
    }
    this.touch();
  }

  setFilterValues(sheetId: string, column: number, allowed: string[] | null): void {
    const sheet = this.requireSheet(sheetId);
    if (!sheet.filter) return;
    this.checkpoint();
    if (allowed == null) delete sheet.filter.criteria[column];
    else sheet.filter.criteria[column] = allowed;
    this.recomputeFilter(sheet);
    this.touch();
  }

  distinctValues(sheetId: string, column: number, range: CellRange): string[] {
    const area = normalizeRange(range);
    const values = new Set<string>();
    const last = Math.min(area.r2, area.r1 + 5000);
    for (let r = area.r1 + 1; r <= last; r += 1) values.add(this.display(sheetId, r, column));
    return [...values].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  }

  addChart(sheetId: string, type: ChartType, range: CellRange, title: string): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    const area = normalizeRange(range);
    sheet.charts.push({
      id: crypto.randomUUID(),
      type,
      range: area,
      title: title || 'Chart',
      x: 36 + sheet.charts.length * 24,
      y: 280 + sheet.charts.length * 16,
      w: 420,
      h: 240,
    });
    this.touch();
  }

  updateChart(sheetId: string, chartId: string, patch: Partial<ChartModel>): void {
    const sheet = this.requireSheet(sheetId);
    const chart = sheet.charts.find((item) => item.id === chartId);
    if (!chart) return;
    Object.assign(chart, patch);
    this.mark();
  }

  removeChart(sheetId: string, chartId: string): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    sheet.charts = sheet.charts.filter((chart) => chart.id !== chartId);
    this.touch();
  }

  addConditional(sheetId: string, rule: ConditionalRule): void {
    this.checkpoint();
    this.requireSheet(sheetId).conditional.push(rule);
    this.touch();
  }

  removeConditional(sheetId: string, ruleId: string): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    sheet.conditional = sheet.conditional.filter((rule) => rule.id !== ruleId);
    this.touch();
  }

  addValidation(sheetId: string, rule: ValidationRule): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    sheet.validations = sheet.validations.filter((item) => !overlaps(item.range, rule.range));
    sheet.validations.push(rule);
    this.touch();
  }

  removeValidation(sheetId: string, ruleId: string): void {
    this.checkpoint();
    const sheet = this.requireSheet(sheetId);
    sheet.validations = sheet.validations.filter((rule) => rule.id !== ruleId);
    this.touch();
  }

  validationAt(sheetId: string, r: number, c: number): ValidationRule | null {
    const sheet = this.sheet(sheetId);
    if (!sheet) return null;
    return sheet.validations.find((rule) => rangeContains(rule.range, r, c)) ?? null;
  }

  defineName(name: string, sheetId: string, range: CellRange): boolean {
    const clean = name.trim().replace(/\s+/g, '_');
    if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(clean)) return false;
    if (!this.sheet(sheetId)) return false;
    this.checkpoint();
    this.model.names = this.model.names.filter((item) => item.name.toLowerCase() !== clean.toLowerCase());
    const named: NamedRange = { name: clean, sheetId, range: normalizeRange(range) };
    this.model.names.push(named);
    this.syncNames();
    this.touch();
    return true;
  }

  removeName(name: string): void {
    this.checkpoint();
    this.model.names = this.model.names.filter((item) => item.name !== name);
    this.syncNames();
    this.touch();
  }

  nameForRange(sheetId: string, range: CellRange): string | null {
    const area = normalizeRange(range);
    const found = this.model.names.find((item) => item.sheetId === sheetId && sameRange(item.range, area));
    return found?.name ?? null;
  }

  addSheet(name?: string): string {
    this.checkpoint();
    const sheet = createSheet(uniqueSheetName(this.model.sheets, sanitizeSheetName(name || 'Sheet')));
    this.model.sheets.push(sheet);
    this.model.activeSheetId = sheet.id;
    this.hf.addSheet(sheet.name);
    this.touch();
    return sheet.id;
  }

  renameSheet(sheetId: string, name: string): boolean {
    const sheet = this.requireSheet(sheetId);
    const clean = sanitizeSheetName(name);
    if (this.model.sheets.some((item) => item.id !== sheetId && item.name.toLowerCase() === clean.toLowerCase())) return false;
    if (clean === sheet.name) return true;
    this.checkpoint();
    this.hf.renameSheet(this.hfId(sheet), clean);
    sheet.name = clean;
    for (const item of this.model.sheets) this.resyncFormulas(item);
    this.syncNames();
    this.touch();
    return true;
  }

  deleteSheet(sheetId: string): boolean {
    if (this.model.sheets.length <= 1) return false;
    this.checkpoint();
    const index = this.model.sheets.findIndex((sheet) => sheet.id === sheetId);
    if (index < 0) return false;
    this.model.sheets.splice(index, 1);
    if (this.model.activeSheetId === sheetId) {
      const next = this.model.sheets[Math.max(0, index - 1)];
      this.model.activeSheetId = next.id;
    }
    this.model.names = this.model.names.filter((item) => item.sheetId !== sheetId);
    this.rebuild();
    this.touch();
    return true;
  }

  moveSheet(sheetId: string, toIndex: number): void {
    const from = this.model.sheets.findIndex((sheet) => sheet.id === sheetId);
    if (from < 0 || from === toIndex) return;
    this.checkpoint();
    const [sheet] = this.model.sheets.splice(from, 1);
    this.model.sheets.splice(Math.max(0, Math.min(toIndex, this.model.sheets.length)), 0, sheet);
    this.touch();
  }

  replaceModel(model: WorkbookModel): void {
    this.model = cloneData(model);
    this.past = [];
    this.future = [];
    this.rebuild();
    this.touch();
  }

  undo(): boolean {
    const prev = this.past.pop();
    if (!prev) return false;
    this.future.push(cloneData(this.model));
    this.model = prev;
    this.rebuild();
    this.touch();
    return true;
  }

  redo(): boolean {
    const next = this.future.pop();
    if (!next) return false;
    this.past.push(cloneData(this.model));
    this.model = next;
    this.rebuild();
    this.touch();
    return true;
  }

  precedents(sheetId: string, r: number, c: number): CellRange[] {
    const sheet = this.sheet(sheetId);
    if (!sheet) return [];
    const list = this.hf.getCellPrecedents({ sheet: this.hfId(sheet), row: r, col: c });
    const ranges: CellRange[] = [];
    for (const item of list) {
      if (item && typeof item === 'object' && 'start' in item && 'end' in item) {
        const start = item.start;
        const end = item.end;
        if (this.sheetIdByHf(start.sheet) !== sheetId) continue;
        ranges.push({ r1: start.row, c1: start.col, r2: end.row, c2: end.col });
      } else if (item && typeof item === 'object' && 'row' in item && 'col' in item) {
        if (this.sheetIdByHf(item.sheet) !== sheetId) continue;
        ranges.push({ r1: item.row, c1: item.col, r2: item.row, c2: item.col });
      }
    }
    return ranges;
  }

  snapshotRange(sheetId: string, range: CellRange): LinkedRangeCell[][] {
    const area = normalizeRange(range);
    const rows: LinkedRangeCell[][] = [];
    for (let r = area.r1; r <= area.r2; r += 1) {
      const row: LinkedRangeCell[] = [];
      for (let c = area.c1; c <= area.c2; c += 1) {
        row.push({
          input: this.inputAt(sheetId, r, c) || null,
          value: this.raw(sheetId, r, c),
          display: this.display(sheetId, r, c),
        });
      }
      rows.push(row);
    }
    return rows;
  }

  isEmpty(sheetId: string, r: number, c: number): boolean {
    return !this.inputAt(sheetId, r, c);
  }

  anchor(sheetId: string, r: number, c: number): { r: number; c: number; range: CellRange } {
    const sheet = this.sheet(sheetId);
    const merge = sheet ? findMerge(sheet, r, c) : null;
    if (merge) return { r: merge.r1, c: merge.c1, range: merge };
    return { r, c, range: { r1: r, c1: c, r2: r, c2: c } };
  }

  private shiftAxis(sheetId: string, axis: 'row' | 'col', at: number, count: number): boolean {
    if (count === 0) return false;
    const sheet = this.requireSheet(sheetId);
    const hfId = this.hfId(sheet);
    this.checkpoint();
    try {
      if (axis === 'row') {
        if (count > 0) this.hf.addRows(hfId, [at, count]);
        else this.hf.removeRows(hfId, [at, -count]);
      } else if (count > 0) this.hf.addColumns(hfId, [at, count]);
      else this.hf.removeColumns(hfId, [at, -count]);
    } catch {
      this.past.pop();
      return false;
    }
    shiftSheetGeometry(sheet, axis, at, count);
    for (const named of this.model.names) {
      if (named.sheetId !== sheet.id) continue;
      const next = shiftRange(named.range, axis, at, count);
      if (next) named.range = next;
    }
    this.model.names = this.model.names.filter((named) => named.sheetId !== sheet.id || named.range.r2 >= named.range.r1);
    this.realignSheet(sheet);
    for (const other of this.model.sheets) {
      if (other.id !== sheet.id) this.resyncFormulas(other);
    }
    this.syncNames();
    this.recomputeFilter(sheet);
    this.touch();
    return true;
  }

  private writeCell(sheet: SheetModel, r: number, c: number, rawInput: string): void {
    const interpreted = interpretInput(rawInput);
    const key = keyOf(r, c);
    const existing = sheet.cells[key];
    const hfId = this.hfId(sheet);
    if (!interpreted.stored) {
      this.hf.setCellContents({ sheet: hfId, row: r, col: c }, null);
      if (existing?.style) sheet.cells[key] = { style: existing.style };
      else delete sheet.cells[key];
      return;
    }
    this.pushEngine(hfId, r, c, interpreted.stored);
    const style: CellStyle = { ...(existing?.style ?? {}) };
    if (interpreted.format) style.format = interpreted.format;
    const record: CellRecord = { input: interpreted.stored };
    if (Object.keys(style).length) record.style = style;
    sheet.cells[key] = record;
  }

  private pushEngine(hfId: number, r: number, c: number, stored: string): void {
    if (!stored) {
      this.hf.setCellContents({ sheet: hfId, row: r, col: c }, null);
      return;
    }
    const value = stored.startsWith('=') ? normalizeFormula(stored) : coerceEngineValue(stored);
    this.hf.setCellContents({ sheet: hfId, row: r, col: c }, value);
  }

  private realignSheet(sheet: SheetModel): void {
    const hfId = this.hfId(sheet);
    for (const key of Object.keys(sheet.cells)) {
      const cell = sheet.cells[key];
      if (!cell) continue;
      const { r, c } = parseKey(key);
      const formula = this.hf.getCellFormula({ sheet: hfId, row: r, col: c });
      if (formula) {
        cell.input = formula.startsWith('=') ? formula : `=${formula}`;
        continue;
      }
      if (!cell.input) continue;
      const value = plainFromUnknown(this.hf.getCellValue({ sheet: hfId, row: r, col: c }));
      if (value == null || value === '') {
        delete cell.input;
        if (!cell.style) delete sheet.cells[key];
      } else if (typeof value === 'number') cell.input = String(value);
      else if (typeof value === 'boolean') cell.input = value ? 'TRUE' : 'FALSE';
      else if (!String(value).startsWith('#')) cell.input = String(value);
    }
  }

  private resyncFormulas(sheet: SheetModel): void {
    const hfId = this.hfId(sheet);
    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (!cell.input?.startsWith('=')) continue;
      const { r, c } = parseKey(key);
      const formula = this.hf.getCellFormula({ sheet: hfId, row: r, col: c });
      if (formula) cell.input = formula.startsWith('=') ? formula : `=${formula}`;
    }
  }

  private recomputeFilter(sheet: SheetModel): void {
    if (!sheet.filter) {
      sheet.filterHidden = [];
      return;
    }
    const hidden: number[] = [];
    const { range, headerRow, criteria } = sheet.filter;
    const entries = Object.entries(criteria);
    if (!entries.length) {
      sheet.filterHidden = [];
      return;
    }
    const last = Math.min(range.r2, range.r1 + 20000);
    for (let r = range.r1; r <= last; r += 1) {
      if (r === headerRow) continue;
      for (const [col, allowed] of entries) {
        if (!allowed) continue;
        const display = this.display(sheet.id, r, Number(col));
        if (!allowed.includes(display)) {
          hidden.push(r);
          break;
        }
      }
    }
    sheet.filterHidden = hidden;
  }

  private syncNames(): void {
    for (const name of [...this.hf.listNamedExpressions()]) {
      try {
        this.hf.removeNamedExpression(name);
      } catch {
        /* The name may already have been removed with its sheet. */
      }
    }
    for (const named of this.model.names) {
      const sheet = this.sheet(named.sheetId);
      if (!sheet) continue;
      const formula = `=${quoteSheet(sheet.name)}!${formatRangeA1(named.range, true)}`;
      try {
        this.hf.addNamedExpression(named.name, formula);
      } catch {
        /* Ignore a name the engine will not accept. The model still keeps it. */
      }
    }
  }

  private hydrate(): void {
    this.hf.batch(() => {
      for (const sheet of this.model.sheets) {
        const existing = this.hf.getSheetId(sheet.name);
        if (existing == null) this.hf.addSheet(sheet.name);
      }
      for (const sheet of this.model.sheets) {
        const hfId = this.hf.getSheetId(sheet.name);
        if (hfId == null) continue;
        for (const [key, cell] of Object.entries(sheet.cells)) {
          if (!cell.input) continue;
          const { r, c } = parseKey(key);
          if (r >= ROW_COUNT || c >= COL_COUNT) continue;
          this.pushEngine(hfId, r, c, cell.input);
        }
      }
    });
    this.syncNames();
  }

  private rebuild(): void {
    this.hf.destroy();
    this.hf = createEngine();
    this.hydrate();
  }

  private checkpoint(): void {
    if (this.gesture) return;
    this.past.push(cloneData(this.model));
    if (this.past.length > 80) this.past.shift();
    this.future = [];
  }

  private touch(): void {
    this.model.updatedAt = Date.now();
    this.version += 1;
    this.dirty = true;
  }

  private mark(): void {
    this.version += 1;
    this.dirty = true;
  }

  private eachCell(sheet: SheetModel, area: CellRange, visit: (r: number, c: number) => void): void {
    const count = (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1);
    if (count > 4000) {
      for (const key of Object.keys(sheet.cells)) {
        const pos = parseKey(key);
        if (rangeContains(area, pos.r, pos.c)) visit(pos.r, pos.c);
      }
      return;
    }
    for (let r = area.r1; r <= area.r2; r += 1) {
      for (let c = area.c1; c <= area.c2; c += 1) visit(r, c);
    }
  }

  private requireSheet(id: string): SheetModel {
    const sheet = this.sheet(id);
    if (!sheet) throw new Error('That sheet is not on the desk.');
    return sheet;
  }

  private hfId(sheet: SheetModel): number {
    const id = this.hf.getSheetId(sheet.name);
    if (id == null) throw new Error(`The engine has no sheet named ${sheet.name}.`);
    return id;
  }

  private sheetIdByHf(hfSheet: number): string | null {
    const name = this.hf.getSheetName(hfSheet);
    if (!name) return null;
    return this.model.sheets.find((sheet) => sheet.name === name)?.id ?? null;
  }
}

function overlaps(a: CellRange, b: CellRange): boolean {
  const x = normalizeRange(a);
  const y = normalizeRange(b);
  return x.r1 <= y.r2 && x.r2 >= y.r1 && x.c1 <= y.c2 && x.c2 >= y.c1;
}

function sameRange(a: CellRange, b: CellRange): boolean {
  const x = normalizeRange(a);
  const y = normalizeRange(b);
  return x.r1 === y.r1 && x.c1 === y.c1 && x.r2 === y.r2 && x.c2 === y.c2;
}
