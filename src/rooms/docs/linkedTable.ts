import type { LinkedRange } from '../../data/linkedRanges';

export interface LinkedTableModel {
  linkId: string;
  caption: string;
  rows: string[][];
}

export function tableFromRange(range: LinkedRange): LinkedTableModel {
  return {
    linkId: range.id,
    caption: `${range.sheetName} · ${range.a1}`,
    rows: range.cells.map((row) => row.map((cell) => cell.display || (cell.value == null ? '' : String(cell.value)))),
  };
}

/** Stable text used to notice when a sheet has new values. */
export function tableSignature(rows: string[][]): string {
  return rows.map((row) => row.join('\u001f')).join('\u001e');
}
