import ExcelJS, {
  type Alignment,
  type Borders,
  type Cell,
  type CellValue,
  type Font,
  type Worksheet,
} from 'exceljs-browser';
import { parseA1Range } from '../domain/a1';
import { dateToExcelSerial, hexForXlsx, hexFromXlsx } from '../domain/format';
import { parseCsv } from '../domain/clipboard';
import {
  createSheet,
  createWorkbook,
  keyOf,
  sanitizeSheetName,
  type BorderEdge,
  type CellStyle,
  type SheetModel,
  type WorkbookModel,
} from '../domain/types';

const FORMAT_TO_NUM: Record<string, string> = {
  number: '#,##0.00',
  currency: '"$"#,##0.00',
  percent: '0.00%',
  date: 'mmm d, yyyy',
};

function numToFormat(fmt: string | undefined): string | undefined {
  if (!fmt) return undefined;
  const text = fmt.toLowerCase();
  if (text.includes('%')) return 'percent';
  if (text.includes('$') || text.includes('¥') || text.includes('€')) return 'currency';
  if (text.includes('yy') || text.includes('dd') || text.includes('mmm')) return 'date';
  if (text.includes('0.00') || text.includes('#,##0')) return 'number';
  return undefined;
}

function excelSheetName(name: string, used: Set<string>): string {
  const base = sanitizeSheetName(name).slice(0, 31) || 'Sheet';
  let next = base;
  let n = 2;
  while (used.has(next.toLowerCase())) {
    const suffix = ` ${n}`;
    next = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    n += 1;
  }
  used.add(next.toLowerCase());
  return next;
}

export async function workbookToXlsx(model: WorkbookModel): Promise<Uint8Array> {
  const book = new ExcelJS.Workbook();
  book.creator = 'AI TEK SOLUTIONS';
  const used = new Set<string>();
  for (const sheet of model.sheets) {
    const ws = book.addWorksheet(excelSheetName(sheet.name, used));
    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (!cell.input && !cell.style) continue;
      const [rowText, colText] = key.split(',');
      const r = Number(rowText);
      const c = Number(colText);
      if (!Number.isFinite(r) || !Number.isFinite(c)) continue;
      const target = ws.getCell(r + 1, c + 1);
      if (cell.input?.startsWith('=')) target.value = { formula: cell.input.slice(1) };
      else if (cell.input === 'TRUE' || cell.input === 'FALSE') target.value = cell.input === 'TRUE';
      else if (cell.input && /^-?\d+(\.\d+)?$/.test(cell.input)) target.value = Number(cell.input);
      else if (cell.input) target.value = cell.input;
      applyStyle(target, cell.style);
    }
    for (const merge of sheet.merges) {
      if (merge.r1 === merge.r2 && merge.c1 === merge.c2) continue;
      ws.mergeCells(merge.r1 + 1, merge.c1 + 1, merge.r2 + 1, merge.c2 + 1);
    }
    for (const [key, width] of Object.entries(sheet.colWidths)) {
      ws.getColumn(Number(key) + 1).width = Math.max(4, width / 8);
    }
    for (const col of sheet.hiddenCols) ws.getColumn(col + 1).hidden = true;
    for (const [key, height] of Object.entries(sheet.rowHeights)) {
      ws.getRow(Number(key) + 1).height = height;
    }
    for (const row of sheet.hiddenRows) ws.getRow(row + 1).hidden = true;
  }
  const out = await book.xlsx.writeBuffer();
  return new Uint8Array(out);
}

function applyStyle(cell: Cell, style: CellStyle | undefined): void {
  if (!style) return;
  const font: Partial<Font> = {};
  if (style.bold) font.bold = true;
  if (style.italic) font.italic = true;
  if (style.underline) font.underline = true;
  if (style.strike) font.strike = true;
  if (style.fontFamily) font.name = style.fontFamily;
  if (style.fontSize) font.size = style.fontSize;
  if (style.color) font.color = { argb: hexForXlsx(style.color) };
  if (Object.keys(font).length) cell.font = font;
  if (style.fill) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexForXlsx(style.fill) } };
  }
  const alignment: Partial<Alignment> = {};
  if (style.align) alignment.horizontal = style.align;
  if (style.valign) alignment.vertical = style.valign;
  if (style.wrap) alignment.wrapText = true;
  if (Object.keys(alignment).length) cell.alignment = alignment;
  if (style.format && FORMAT_TO_NUM[style.format]) cell.numFmt = FORMAT_TO_NUM[style.format];
  if (style.border) {
    const border: Partial<Borders> = {};
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const edge = style.border[side];
      if (!edge) continue;
      border[side] = { style: edge.style, color: { argb: hexForXlsx(edge.color) } };
    }
    cell.border = border;
  }
}

export async function workbookFromXlsx(data: ArrayBuffer | Uint8Array, title: string): Promise<WorkbookModel> {
  const book = new ExcelJS.Workbook();
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  await book.xlsx.load(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer);
  const model = createWorkbook(title.replace(/\.xlsx$/i, '') || 'Workbook');
  model.sheets = [];
  for (const ws of book.worksheets) {
    model.sheets.push(sheetFromWorksheet(ws));
  }
  if (!model.sheets.length) model.sheets.push(createSheet('Sheet1'));
  model.activeSheetId = model.sheets[0].id;
  return model;
}

function sheetFromWorksheet(ws: Worksheet): SheetModel {
  const sheet = createSheet(sanitizeSheetName(ws.name));
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const record = recordFromCell(cell);
      if (!record) return;
      sheet.cells[keyOf(rowNumber - 1, colNumber - 1)] = record;
    });
    if (row.hidden) sheet.hiddenRows.push(rowNumber - 1);
    if (row.height && Math.abs(row.height - 15) > 1) sheet.rowHeights[rowNumber - 1] = Math.round(row.height);
  });
  for (const ref of ws.model.merges ?? []) {
    const range = parseA1Range(ref.replace(/\$/g, ''));
    if (range) sheet.merges.push(range);
  }
  const colCount = ws.columnCount || 0;
  for (let c = 1; c <= colCount; c += 1) {
    const column = ws.getColumn(c);
    if (column.hidden) sheet.hiddenCols.push(c - 1);
    if (typeof column.width === 'number' && column.width > 0) {
      sheet.colWidths[c - 1] = Math.round(column.width * 8);
    }
  }
  return sheet;
}

function recordFromCell(cell: Cell): { input?: string; style?: CellStyle } | null {
  const input = inputFromValue(cell.value);
  const style = styleFromCell(cell);
  if (!input && !style) return null;
  return { input, style };
}

function inputFromValue(value: CellValue): string | undefined {
  if (value == null) return undefined;
  if (typeof value === 'number') return String(value);
  if (typeof value === 'string') return value;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (value instanceof Date) return String(dateToExcelSerial(value));
  if (typeof value === 'object') {
    if ('formula' in value && typeof value.formula === 'string') return `=${value.formula}`;
    if ('richText' in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text).join('');
    }
    if ('text' in value && typeof value.text === 'string') return value.text;
    if ('error' in value && typeof value.error === 'string') return value.error;
  }
  return undefined;
}

function styleFromCell(cell: Cell): CellStyle | undefined {
  const style: CellStyle = {};
  const font = cell.font;
  if (font?.bold) style.bold = true;
  if (font?.italic) style.italic = true;
  if (font?.underline) style.underline = true;
  if (font?.strike) style.strike = true;
  if (font?.name && font.name !== 'Calibri') style.fontFamily = font.name;
  if (font?.size && font.size !== 11) style.fontSize = font.size;
  const color = hexFromXlsx(font?.color?.argb);
  if (color && color.toLowerCase() !== '#000000') style.color = color;
  const fill = cell.fill;
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid' && fill.fgColor && 'argb' in fill.fgColor) {
    const hex = hexFromXlsx(fill.fgColor.argb);
    if (hex) style.fill = hex;
  }
  if (cell.alignment?.horizontal === 'left' || cell.alignment?.horizontal === 'center' || cell.alignment?.horizontal === 'right') {
    style.align = cell.alignment.horizontal;
  }
  if (cell.alignment?.vertical === 'top' || cell.alignment?.vertical === 'middle' || cell.alignment?.vertical === 'bottom') {
    style.valign = cell.alignment.vertical;
  }
  if (cell.alignment?.wrapText) style.wrap = true;
  const format = numToFormat(cell.numFmt);
  if (format) style.format = format;
  if (cell.border) {
    const border: NonNullable<CellStyle['border']> = {};
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const edge = cell.border[side];
      if (!edge?.style) continue;
      const weight: BorderEdge['style'] = edge.style === 'medium' || edge.style === 'thick' ? edge.style : 'thin';
      const edgeColor = edge.color && 'argb' in edge.color ? hexFromXlsx(edge.color.argb) : undefined;
      border[side] = { style: weight, color: edgeColor ?? '#1c2430' };
    }
    if (Object.keys(border).length) style.border = border;
  }
  return Object.keys(style).length ? style : undefined;
}

export function workbookFromCsv(name: string, text: string): WorkbookModel {
  const rows = parseCsv(text);
  const model = createWorkbook(name.replace(/\.csv$/i, '') || 'Import');
  const sheet = model.sheets[0];
  rows.forEach((row, r) => {
    row.forEach((value, c) => {
      if (value === '') return;
      sheet.cells[keyOf(r, c)] = { input: value };
    });
  });
  return model;
}

export function sheetToDisplayRows(model: SheetModel, display: (r: number, c: number) => string): string[][] {
  let maxR = 0;
  let maxC = 0;
  let any = false;
  for (const key of Object.keys(model.cells)) {
    if (!model.cells[key]?.input) continue;
    const [rowText, colText] = key.split(',');
    const r = Number(rowText);
    const c = Number(colText);
    any = true;
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  }
  if (!any) return [];
  const rows: string[][] = [];
  for (let r = 0; r <= maxR; r += 1) {
    const row: string[] = [];
    for (let c = 0; c <= maxC; c += 1) row.push(display(r, c));
    rows.push(row);
  }
  return rows;
}
