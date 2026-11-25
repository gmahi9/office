import { describe, expect, it } from 'vitest';
import type { LinkedRange } from '../../data/linkedRanges';
import { tableFromRange, tableSignature } from './linkedTable';

function range(display: string[][]): LinkedRange {
  return {
    id: 'link_desk_pipeline',
    workbookId: 'book_desk_plan',
    sheetId: 'sheet_pipeline',
    sheetName: 'Pipeline',
    a1: 'A4:B5',
    r1: 3,
    c1: 0,
    r2: 4,
    c2: 1,
    cells: display.map((row) => row.map((cell) => ({ input: cell, value: cell, display: cell }))),
    createdAt: 1,
    updatedAt: 1,
  };
}

describe('linked tables', () => {
  it('builds rows from a sheet snapshot and notices when values change', () => {
    const first = tableFromRange(range([
      ['Role', 'Openings'],
      ['Analyst', '4'],
    ]));
    const next = tableFromRange(range([
      ['Role', 'Openings'],
      ['Analyst', '5'],
    ]));
    expect(first.caption).toBe('Pipeline · A4:B5');
    expect(first.rows).toEqual([['Role', 'Openings'], ['Analyst', '4']]);
    expect(tableSignature(first.rows)).not.toBe(tableSignature(next.rows));
    expect(next.rows[1]?.[1]).toBe('5');
  });
});
