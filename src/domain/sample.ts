import type { WorkbookModel } from './types';
import { createSheet } from './types';

const PIPELINE = 'sheet_pipeline';
const RATES = 'sheet_rates';

/** A desk plan for the staffing practice, so a new browser is not an empty grid. */
export function createSampleWorkbook(x = 28, y = 64): WorkbookModel {
  const pipeline = createSheet('Pipeline', PIPELINE);
  const rates = createSheet('Rates', RATES);
  const now = Date.now();

  const title = (text: string) => ({ input: text, style: { bold: true, fontFamily: 'Syne', fontSize: 16 } });
  pipeline.cells['0,0'] = title('Placement desk');
  pipeline.cells['1,0'] = {
    input: 'AI TEK SOLUTIONS · roles in motion',
    style: { italic: true, fontSize: 12 },
  };
  pipeline.merges.push({ r1: 0, c1: 0, r2: 0, c2: 7 });
  pipeline.merges.push({ r1: 1, c1: 0, r2: 1, c2: 7 });

  const headers = ['Role', 'Practice', 'City', 'Openings', 'Filled', 'Remaining', 'Rate', 'Week value'];
  headers.forEach((label, c) => {
    pipeline.cells[`3,${c}`] = {
      input: label,
      style: {
        bold: true,
        fill: '#e4eef6',
        color: '#1b3148',
        align: c >= 3 ? 'right' : 'left',
        border: { bottom: { style: 'thin', color: '#8d4520' } },
      },
    };
  });

  const roles: Array<[string, string, string, number, number]> = [
    ['Cybersecurity analyst', 'Technology', 'Austin', 4, 2],
    ['Staff nurse', 'Healthcare', 'Denver', 6, 6],
    ['CNC machinist', 'Skilled trades', 'Toledo', 3, 1],
    ['Controller', 'Finance', 'Chicago', 1, 0],
    ['Warehouse lead', 'Logistics', 'Dallas', 5, 3],
    ['Maintenance tech', 'Skilled trades', 'Tulsa', 2, 2],
  ];
  roles.forEach((role, i) => {
    const r = 4 + i;
    const row = r + 1;
    pipeline.cells[`${r},0`] = { input: role[0], style: { align: 'left' } };
    pipeline.cells[`${r},1`] = { input: role[1] };
    pipeline.cells[`${r},2`] = { input: role[2] };
    pipeline.cells[`${r},3`] = { input: String(role[3]), style: { align: 'right' } };
    pipeline.cells[`${r},4`] = { input: String(role[4]), style: { align: 'right' } };
    pipeline.cells[`${r},5`] = { input: `=D${row}-E${row}`, style: { align: 'right' } };
    pipeline.cells[`${r},6`] = { input: `=VLOOKUP(B${row},RateCard,2,FALSE)`, style: { align: 'right', format: 'currency' } };
    pipeline.cells[`${r},7`] = { input: `=F${row}*G${row}*40`, style: { align: 'right', format: 'currency' } };
  });

  pipeline.cells['10,0'] = { input: 'Desk total', style: { bold: true } };
  pipeline.cells['10,3'] = { input: '=SUM(D5:D10)', style: { bold: true, align: 'right' } };
  pipeline.cells['10,4'] = { input: '=SUM(E5:E10)', style: { bold: true, align: 'right' } };
  pipeline.cells['10,5'] = { input: '=SUM(F5:F10)', style: { bold: true, align: 'right' } };
  pipeline.cells['10,7'] = { input: '=SUM(H5:H10)', style: { bold: true, align: 'right', format: 'currency' } };

  pipeline.conditional.push({
    id: 'open-roles',
    range: { r1: 4, c1: 5, r2: 9, c2: 5 },
    type: 'gt',
    value: '0',
    fill: '#f3e1d4',
    color: '#8d4520',
  });
  pipeline.conditional.push({
    id: 'filled-roles',
    range: { r1: 4, c1: 5, r2: 9, c2: 5 },
    type: 'eq',
    value: '0',
    fill: '#d7ebf8',
    color: '#1b3148',
  });
  pipeline.validations.push({
    id: 'practice-list',
    range: { r1: 4, c1: 1, r2: 9, c2: 1 },
    type: 'list',
    values: ['Technology', 'Healthcare', 'Finance', 'Skilled trades', 'Logistics'],
  });
  pipeline.freeze = { rows: 4, cols: 0 };
  pipeline.colWidths = { 0: 196, 1: 130, 2: 110, 3: 100, 4: 90, 5: 110, 6: 100, 7: 120 };
  pipeline.rowHeights = { 0: 32 };
  pipeline.charts.push({
    id: 'openings-chart',
    type: 'bar',
    range: { r1: 3, c1: 0, r2: 9, c2: 3 },
    title: 'Openings by role',
    x: 24,
    y: 348,
    w: 460,
    h: 240,
  });

  rates.cells['0,0'] = { input: 'Practice', style: { bold: true, fill: '#e4eef6' } };
  rates.cells['0,1'] = { input: 'Standard rate', style: { bold: true, fill: '#e4eef6', align: 'right' } };
  const rateRows: Array<[string, number]> = [
    ['Technology', 96],
    ['Healthcare', 74],
    ['Finance', 88],
    ['Skilled trades', 62],
    ['Logistics', 54],
  ];
  rateRows.forEach((row, i) => {
    rates.cells[`${i + 1},0`] = { input: row[0] };
    rates.cells[`${i + 1},1`] = { input: String(row[1]), style: { format: 'currency', align: 'right' } };
  });
  rates.colWidths = { 0: 160, 1: 130 };
  rates.cells['7,0'] = {
    input: 'RateCard is the named range A2:B6. Pipeline looks it up.',
    style: { italic: true, color: '#5e6b7c' },
  };

  return {
    id: 'book_desk_plan',
    title: 'Desk plan',
    createdAt: now,
    updatedAt: now,
    activeSheetId: pipeline.id,
    sheets: [pipeline, rates],
    names: [
      {
        name: 'RateCard',
        sheetId: rates.id,
        range: { r1: 1, c1: 0, r2: 5, c2: 1 },
      },
    ],
    x,
    y,
  };
}
