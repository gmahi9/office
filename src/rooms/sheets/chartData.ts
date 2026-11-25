import { colToA } from '../../domain/a1';
import { normalizeRange, type CellRange } from '../../domain/types';
import type { WorkbookSession } from '../../engine/session';

export interface ChartSeries {
  name: string;
  values: Array<number | null>;
}

export interface ChartData {
  categories: string[];
  series: ChartSeries[];
  xs: number[];
}

export function readChartData(session: WorkbookSession, sheetId: string, range: CellRange): ChartData {
  const area = normalizeRange(range);
  const categories: string[] = [];
  const xs: number[] = [];
  const firstRaw = session.raw(sheetId, area.r1 + 1, area.c1);
  const firstIsNumber = typeof firstRaw === 'number';
  for (let r = area.r1 + 1; r <= area.r2 && categories.length < 40; r += 1) {
    const label = session.display(sheetId, r, area.c1);
    categories.push(label || `${r + 1}`);
    const x = session.raw(sheetId, r, area.c1);
    xs.push(typeof x === 'number' ? x : categories.length);
  }
  const series: ChartSeries[] = [];
  const startCol = firstIsNumber ? area.c1 : area.c1 + 1;
  for (let c = startCol; c <= area.c2; c += 1) {
    if (!firstIsNumber && c === area.c1) continue;
    const values: Array<number | null> = [];
    let numeric = 0;
    for (let r = area.r1 + 1; r <= area.r2 && values.length < 40; r += 1) {
      const raw = session.raw(sheetId, r, c);
      if (typeof raw === 'number') {
        values.push(raw);
        numeric += 1;
      } else values.push(null);
    }
    if (!numeric) continue;
    const header = session.display(sheetId, area.r1, c);
    series.push({ name: header || colToA(c), values });
  }
  return { categories, series, xs };
}

export const CHART_COLORS = ['#2f7eb8', '#8d4520', '#3d93c9', '#1b3148', '#c47a45', '#5e6b7c'];
