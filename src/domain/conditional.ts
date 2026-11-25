import type { ConditionalRule, PlainValue, SheetModel } from './types';
import { parseKey, rangeContains } from './types';
import { mixHex } from './format';

export interface Paint {
  fill?: string;
  color?: string;
}

export function ruleMatches(rule: ConditionalRule, value: PlainValue, scaleT?: number): Paint | null {
  if (rule.type === 'colorScale') {
    if (typeof value !== 'number' || scaleT == null) return null;
    const min = rule.minColor || '#c5dff3';
    const max = rule.maxColor || '#8d4520';
    return { fill: mixHex(min, max, scaleT) };
  }
  const hit = compare(rule, value);
  if (!hit) return null;
  const paint: Paint = {};
  if (rule.fill) paint.fill = rule.fill;
  if (rule.color) paint.color = rule.color;
  return paint.fill || paint.color ? paint : null;
}

function compare(rule: ConditionalRule, value: PlainValue): boolean {
  if (value == null || value === '') return false;
  if (rule.type === 'contains') {
    return String(value).toLowerCase().includes((rule.value ?? '').toLowerCase());
  }
  const left = typeof value === 'number' ? value : Number(value);
  const right = Number(rule.value);
  const numeric = typeof value === 'number' || (typeof value === 'string' && value !== '' && !Number.isNaN(left));
  if (rule.type === 'between') {
    const high = Number(rule.value2);
    return numeric && !Number.isNaN(right) && !Number.isNaN(high) && left >= Math.min(right, high) && left <= Math.max(right, high);
  }
  if (rule.type === 'eq' && !numeric) {
    return String(value).toLowerCase() === (rule.value ?? '').toLowerCase();
  }
  if (!numeric || Number.isNaN(right)) return false;
  if (rule.type === 'gt') return left > right;
  if (rule.type === 'lt') return left < right;
  if (rule.type === 'eq') return left === right;
  return false;
}

export function scaleBounds(
  sheet: SheetModel,
  rule: ConditionalRule,
  read: (r: number, c: number) => PlainValue,
): { min: number; max: number } {
  let min = Infinity;
  let max = -Infinity;
  for (const key of Object.keys(sheet.cells)) {
    const pos = parseKey(key);
    if (!rangeContains(rule.range, pos.r, pos.c)) continue;
    const value = read(pos.r, pos.c);
    if (typeof value !== 'number') continue;
    if (value < min) min = value;
    if (value > max) max = value;
  }
  if (!Number.isFinite(min) || !Number.isFinite(max)) return { min: 0, max: 1 };
  if (min === max) return { min, max: min + 1 };
  return { min, max };
}

export function paintForCell(
  sheet: SheetModel,
  r: number,
  c: number,
  value: PlainValue,
  bounds: Map<string, { min: number; max: number }>,
): Paint {
  const paint: Paint = {};
  for (const rule of sheet.conditional) {
    if (!rangeContains(rule.range, r, c)) continue;
    let t: number | undefined;
    if (rule.type === 'colorScale' && typeof value === 'number') {
      const span = bounds.get(rule.id) ?? { min: 0, max: 1 };
      t = (value - span.min) / (span.max - span.min);
    }
    const next = ruleMatches(rule, value, t);
    if (!next) continue;
    if (next.fill) paint.fill = next.fill;
    if (next.color) paint.color = next.color;
  }
  return paint;
}
