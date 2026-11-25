import { describe, expect, it } from 'vitest';
import { computeFill } from './fill';
import { adjustFormulaForStructure, normalizeFormula, offsetFormula } from './formula';
import { createSheet } from './types';
import { SizeIndex } from './sizeIndex';
import { jumpCell } from './navigate';

describe('formulas', () => {
  it('offsets relative refs and keeps locks', () => {
    expect(offsetFormula('=A1+B2', 1, 0)).toBe('=A2+B3');
    expect(offsetFormula('=$A$1', 2, 3)).toBe('=$A$1');
    expect(offsetFormula('=Sheet2!B2', 0, 1)).toBe('=Sheet2!C2');
    expect(offsetFormula("='Q4 Plan'!$B3", 0, 1)).toBe("='Q4 Plan'!$B3");
    expect(offsetFormula('=A$1', 4, 1)).toBe('=B$1');
    expect(offsetFormula('=SUM(A1:A10)', 0, 1)).toBe('=SUM(B1:B10)');
  });

  it('leaves text and names alone', () => {
    expect(offsetFormula('=A1&"A1"', 1, 0)).toBe('=A2&"A1"');
    expect(offsetFormula('=RateCard', 1, 1)).toBe('=RateCard');
  });

  it('adjusts refs when rows are inserted or removed', () => {
    expect(adjustFormulaForStructure('=A1+Rates!B2', 'Pipeline', 'Pipeline', 'row', 0, 1)).toBe('=A2+Rates!B2');
    expect(adjustFormulaForStructure('=SUM(A1:A10)', 'Pipeline', 'Pipeline', 'row', 4, 1)).toBe('=SUM(A1:A11)');
    expect(adjustFormulaForStructure('=Rates!B2', 'Pipeline', 'Rates', 'row', 0, 1)).toBe('=Rates!B3');
    expect(adjustFormulaForStructure('=A2', 'Pipeline', 'Pipeline', 'row', 1, -1)).toBe('=#REF!');
    expect(adjustFormulaForStructure('=$A$5', 'Pipeline', 'Pipeline', 'row', 1, 2)).toBe('=$A$7');
  });

  it('rewrites boolean literals for the engine', () => {
    expect(normalizeFormula('=VLOOKUP(B5,RateCard,2,FALSE)')).toBe('=VLOOKUP(B5,RateCard,2,FALSE())');
    expect(normalizeFormula('=AND(TRUE,FALSE())')).toBe('=AND(TRUE(),FALSE())');
    expect(normalizeFormula('="TRUE"')).toBe('="TRUE"');
  });
});

describe('fill', () => {
  it('continues a numeric series and offsets formulas', () => {
    const sheet = createSheet('Sheet1');
    sheet.cells['0,0'] = { input: '1' };
    sheet.cells['1,0'] = { input: '2' };
    sheet.cells['0,1'] = { input: '=A1' };
    const edits = computeFill(sheet, { r1: 0, c1: 0, r2: 1, c2: 1 }, { r1: 0, c1: 0, r2: 3, c2: 1 });
    const byKey = new Map(edits.map((edit) => [`${edit.r},${edit.c}`, edit.record?.input]));
    expect(byKey.get('2,0')).toBe('3');
    expect(byKey.get('3,0')).toBe('4');
    expect(byKey.get('2,1')).toBe('=A3');
  });
});

describe('size index', () => {
  it('skips hidden rows and reaches the last row', () => {
    const index = new SizeIndex(100_000, 26, [1], { 2: 40 });
    expect(index.size(1)).toBe(0);
    expect(index.size(2)).toBe(40);
    expect(index.offset(3)).toBe(26 + 40);
    expect(index.indexAt(index.offset(99_999))).toBe(99_999);
    expect(index.total()).toBe(26 * 99_998 + 40);
  });
});

describe('jump', () => {
  it('lands on the edge of a block', () => {
    const filled = new Set(['0,0', '1,0', '2,0', '5,0']);
    const next = jumpCell(
      (r, c) => !filled.has(`${r},${c}`),
      () => false,
      0,
      0,
      1,
      0,
    );
    expect(next).toEqual({ r: 2, c: 0 });
  });
});
