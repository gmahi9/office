declare module 'exceljs-browser' {
  export interface Font {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean | object;
    strike?: boolean;
    name?: string;
    size?: number;
    color?: { argb?: string };
  }

  export interface Alignment {
    horizontal?: 'left' | 'center' | 'right' | 'justify';
    vertical?: 'top' | 'middle' | 'bottom';
    wrapText?: boolean;
  }

  export interface BorderEdge {
    style?: string;
    color?: { argb?: string };
  }

  export interface Borders {
    top?: BorderEdge;
    right?: BorderEdge;
    bottom?: BorderEdge;
    left?: BorderEdge;
  }

  export interface Fill {
    type?: string;
    pattern?: string;
    fgColor?: { argb?: string };
  }

  export type CellValue =
    | string
    | number
    | boolean
    | Date
    | null
    | undefined
    | { formula?: string; richText?: Array<{ text: string }>; text?: string; error?: string };

  export interface Cell {
    value: CellValue;
    font?: Font;
    fill?: Fill;
    alignment?: Alignment;
    numFmt?: string;
    border?: Borders;
  }

  export interface Row {
    hidden?: boolean;
    height?: number;
    eachCell(opts: { includeEmpty: boolean }, fn: (cell: Cell, col: number) => void): void;
  }

  export interface Column {
    hidden?: boolean;
    width?: number;
  }

  export interface Worksheet {
    name: string;
    columnCount?: number;
    model: { merges?: string[] };
    getCell(row: number, col: number): Cell;
    mergeCells(r1: number, c1: number, r2: number, c2: number): void;
    getColumn(index: number): Column;
    getRow(index: number): Row;
    eachRow(opts: { includeEmpty: boolean }, fn: (row: Row, rowNumber: number) => void): void;
  }

  export class Workbook {
    creator: string;
    worksheets: Worksheet[];
    addWorksheet(name: string): Worksheet;
    xlsx: {
      writeBuffer(): Promise<ArrayBuffer>;
      load(data: ArrayBuffer): Promise<void>;
    };
  }

  const ExcelJS: { Workbook: typeof Workbook };
  export default ExcelJS;
}
