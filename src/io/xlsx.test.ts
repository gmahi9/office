import { describe, expect, it } from 'vitest';
import { createSheet, createWorkbook, keyOf } from '../domain/types';
import { WorkbookSession } from '../engine/session';
import { workbookFromXlsx, workbookToXlsx } from './xlsx';

describe('xlsx round trip', () => {
  it('keeps values, formulas, styles, merges, and sheets', async () => {
    const model = createWorkbook('Round trip');
    const first = model.sheets[0];
    first.name = 'Pipeline';
    first.cells[keyOf(0, 0)] = {
      input: 'Hello',
      style: { bold: true, fill: '#e7c3a8', color: '#1c2430', align: 'center' },
    };
    first.cells[keyOf(1, 0)] = { input: '=A1&"!"' };
    first.cells[keyOf(2, 1)] = { input: '96', style: { format: 'currency' } };
    first.merges.push({ r1: 0, c1: 0, r2: 0, c2: 1 });
    const rates = createSheet('Rates');
    rates.cells[keyOf(0, 1)] = { input: '12' };
    model.sheets.push(rates);

    const bytes = await workbookToXlsx(model);
    const back = await workbookFromXlsx(bytes, 'Round trip');
    expect(back.sheets.map((sheet) => sheet.name)).toEqual(['Pipeline', 'Rates']);
    const pipeline = back.sheets[0];
    expect(pipeline.cells[keyOf(0, 0)]?.input).toBe('Hello');
    expect(pipeline.cells[keyOf(0, 0)]?.style?.bold).toBe(true);
    expect(pipeline.cells[keyOf(0, 0)]?.style?.fill?.toLowerCase()).toBe('#e7c3a8');
    expect(pipeline.cells[keyOf(1, 0)]?.input).toBe('=A1&"!"');
    expect(pipeline.merges).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 1 }]);
    expect(pipeline.cells[keyOf(2, 1)]?.style?.format).toBe('currency');
    expect(back.sheets[1].cells[keyOf(0, 1)]?.input).toBe('12');
  });
});

describe('workbook engine', () => {
  it('evaluates cross-sheet formulas, named ranges, and undo', () => {
    const model = createWorkbook('Calc');
    const first = model.sheets[0];
    first.name = 'Pipeline';
    const rates = createSheet('Rates');
    rates.cells[keyOf(1, 0)] = { input: 'Technology' };
    rates.cells[keyOf(1, 1)] = { input: '96' };
    model.sheets.push(rates);
    model.names.push({
      name: 'RateCard',
      sheetId: rates.id,
      range: { r1: 1, c1: 0, r2: 1, c2: 1 },
    });
    first.cells[keyOf(4, 1)] = { input: 'Technology' };
    first.cells[keyOf(4, 3)] = { input: '4' };
    first.cells[keyOf(4, 4)] = { input: '2' };
    first.cells[keyOf(4, 5)] = { input: '=D5-E5' };
    first.cells[keyOf(4, 6)] = { input: '=VLOOKUP(B5,RateCard,2,FALSE)', style: { format: 'currency' } };

    const session = new WorkbookSession(model);
    expect(session.display(first.id, 4, 5)).toBe('2');
    expect(session.display(first.id, 4, 6)).toBe('$96.00');
    session.setCell(first.id, 2, 0, '=1/0');
    expect(session.raw(first.id, 2, 0)).toBe('#DIV/0!');
    const precedents = session.precedents(first.id, 4, 5);
    expect(precedents.some((range) => range.c1 === 3 && range.r1 === 4)).toBe(true);
    expect(session.undo()).toBe(true);
    expect(session.inputAt(first.id, 2, 0)).toBe('');
    session.destroy();
  });
});
