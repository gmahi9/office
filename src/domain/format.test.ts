import { describe, expect, it } from 'vitest';
import { dateToExcelSerial, formatValue, interpretInput } from './format';

describe('format', () => {
  it('formats currency, percent, and dates', () => {
    expect(formatValue(1200.5, 'currency')).toBe('$1,200.50');
    expect(formatValue(0.25, 'percent')).toBe('25.00%');
    expect(formatValue(dateToExcelSerial(new Date(Date.UTC(2026, 2, 1))), 'date')).toBe('Mar 1, 2026');
    expect(formatValue('#DIV/0!', 'currency')).toBe('#DIV/0!');
  });

  it('interprets typed money and percents', () => {
    expect(interpretInput('$1,200')).toEqual({ stored: '1200', format: 'currency' });
    expect(interpretInput('25%')).toEqual({ stored: '0.25', format: 'percent' });
    expect(interpretInput('=A1+1').stored).toBe('=A1+1');
    expect(interpretInput('2026-03-01').format).toBe('date');
  });
});
