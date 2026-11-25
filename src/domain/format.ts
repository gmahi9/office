import type { PlainValue } from './types';

export function excelSerialToDate(serial: number): Date {
  const whole = Math.floor(serial);
  const frac = serial - whole;
  const ms = Date.UTC(1899, 11, 30) + whole * 86_400_000 + Math.round(frac * 86_400_000);
  return new Date(ms);
}

export function dateToExcelSerial(date: Date): number {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.round((utc - Date.UTC(1899, 11, 30)) / 86_400_000);
}

function formatDate(date: Date): string {
  return date.toLocaleDateString('en-US', {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
    timeZone: 'UTC',
  });
}

function generalNumber(n: number): string {
  if (!Number.isFinite(n)) return '#NUM!';
  if (Object.is(n, -0)) return '0';
  const abs = Math.abs(n);
  if (abs !== 0 && (abs >= 1e12 || abs < 1e-8)) return n.toExponential(2);
  if (Number.isInteger(n)) return n.toLocaleString('en-US');
  return n.toLocaleString('en-US', { maximumFractionDigits: 10 });
}

export function formatValue(value: PlainValue, format?: string): string {
  if (value == null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'string') {
    if (value.startsWith('#')) return value;
    if (format === 'date') {
      const parsed = Date.parse(value);
      if (!Number.isNaN(parsed)) return formatDate(new Date(parsed));
    }
    return value;
  }
  return formatNumber(value, format ?? 'general');
}

export function formatNumber(n: number, format: string): string {
  if (!Number.isFinite(n)) return '#NUM!';
  if (format === 'number') {
    return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  if (format === 'currency') {
    return n.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
  }
  if (format === 'percent') {
    return `${(n * 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}%`;
  }
  if (format === 'date') return formatDate(excelSerialToDate(n));
  return generalNumber(n);
}

export interface InterpretedInput {
  stored: string;
  format?: string;
}

/** Turn a typed literal into something formulas can add. Formulas pass through. */
export function interpretInput(raw: string): InterpretedInput {
  const text = raw.replace(/\r\n/g, '\n');
  if (text.startsWith('=')) return { stored: text };
  const trimmed = text.trim();
  if (trimmed === '') return { stored: '' };
  if (/^-?\d+(\.\d+)?%$/.test(trimmed)) {
    return { stored: String(Number(trimmed.slice(0, -1)) / 100), format: 'percent' };
  }
  if (/^\$\s?-?\d{1,3}(,\d{3})*(\.\d+)?$/.test(trimmed) || /^\$\s?-?\d+(\.\d+)?$/.test(trimmed)) {
    return { stored: trimmed.replace(/[$,\s]/g, ''), format: 'currency' };
  }
  if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(trimmed)) {
    return { stored: trimmed.replace(/,/g, '') };
  }
  if (/^-?\d+(\.\d+)?$/.test(trimmed)) return { stored: trimmed };
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
    const date = new Date(`${trimmed}T00:00:00Z`);
    if (!Number.isNaN(date.getTime())) {
      return { stored: String(dateToExcelSerial(date)), format: 'date' };
    }
  }
  return { stored: text };
}

export function coerceEngineValue(stored: string): string | number | boolean | null {
  if (stored === '') return null;
  if (stored.startsWith('=')) return stored;
  if (stored === 'TRUE') return true;
  if (stored === 'FALSE') return false;
  if (/^-?\d+(\.\d+)?$/.test(stored)) {
    const n = Number(stored);
    if (Number.isFinite(n)) return n;
  }
  return stored;
}

export function plainFromUnknown(value: unknown): PlainValue {
  if (value == null) return null;
  if (typeof value === 'number') return Number.isFinite(value) ? value : '#NUM!';
  if (typeof value === 'string' || typeof value === 'boolean') return value;
  if (typeof value === 'object' && 'value' in value) {
    const inner = (value as { value: unknown }).value;
    if (typeof inner === 'string') return inner;
  }
  return String(value);
}

export function mixHex(from: string, to: string, t: number): string {
  const a = rgb(from);
  const b = rgb(to);
  const mix = (x: number, y: number) => Math.round(x + (y - x) * clamp01(t));
  return `#${[mix(a[0], b[0]), mix(a[1], b[1]), mix(a[2], b[2])]
    .map((n) => n.toString(16).padStart(2, '0'))
    .join('')}`;
}

function clamp01(n: number): number {
  if (n < 0) return 0;
  if (n > 1) return 1;
  return n;
}

function rgb(hex: string): [number, number, number] {
  const h = hex.replace('#', '');
  const full = h.length === 3 ? h.split('').map((c) => c + c).join('') : h.slice(0, 6);
  return [parseInt(full.slice(0, 2), 16), parseInt(full.slice(2, 4), 16), parseInt(full.slice(4, 6), 16)];
}

export function hexForXlsx(color: string): string {
  const h = color.replace('#', '');
  const full = h.length === 3 ? h.split('').map((c) => c + c).join('') : h.slice(0, 6);
  return `FF${full.toUpperCase()}`;
}

export function hexFromXlsx(rgbValue: string | undefined): string | undefined {
  if (!rgbValue) return undefined;
  const hex = rgbValue.replace('#', '');
  if (hex.length === 8) return `#${hex.slice(2)}`;
  if (hex.length === 6) return `#${hex}`;
  return undefined;
}
