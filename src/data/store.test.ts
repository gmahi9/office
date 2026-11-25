import { beforeEach, describe, expect, it } from 'vitest';
import { createWorkbook } from '../domain/types';
import { resetOfficeDb } from './db';
import { deleteLinkedRange, getLinkedRange, listLinkedRanges, saveLinkedRange, type LinkedRange } from './linkedRanges';
import { deleteWorkbook, getWorkbook, listWorkbooks, saveWorkbook, resetDeskCache } from './workbooks';

describe('workbook and linked range store', () => {
  beforeEach(async () => {
    resetDeskCache();
    await resetOfficeDb();
  });

  it('saves a workbook and reads it back', async () => {
    const model = createWorkbook('Pipeline');
    model.sheets[0].cells['0,0'] = { input: 'Role' };
    await saveWorkbook(model);
    const loaded = await getWorkbook(model.id);
    expect(loaded?.title).toBe('Pipeline');
    expect(loaded?.sheets[0].cells['0,0']?.input).toBe('Role');
    expect(await listWorkbooks()).toHaveLength(1);
    await deleteWorkbook(model.id);
    expect(await listWorkbooks()).toHaveLength(0);
  });

  it('publishes a linked range for another room to read', async () => {
    const range: LinkedRange = {
      id: 'link_rate',
      workbookId: 'book_desk_plan',
      sheetId: 'sheet_pipeline',
      sheetName: 'Pipeline',
      a1: 'G5:H10',
      r1: 4,
      c1: 6,
      r2: 9,
      c2: 7,
      cells: [[{ input: '=VLOOKUP(B5,RateCard,2,FALSE)', value: 96, display: '$96.00' }]],
      createdAt: 10,
      updatedAt: 10,
    };
    await saveLinkedRange(range);
    const other: LinkedRange = { ...range, id: 'link_other', workbookId: 'book_other' };
    await saveLinkedRange(other);
    expect(await listLinkedRanges('book_desk_plan')).toHaveLength(1);
    expect((await getLinkedRange('link_rate'))?.cells[0][0].value).toBe(96);
    await deleteLinkedRange('link_rate');
    expect(await getLinkedRange('link_rate')).toBeUndefined();
    expect(await listLinkedRanges()).toHaveLength(1);
  });
});
