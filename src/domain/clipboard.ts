/** Parse clipboard text the way a spreadsheet receives it from Excel or Google Sheets. */

export function parseTsv(text: string): string[][] {
  const normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const trimmed = normalized.endsWith('\n') ? normalized.slice(0, -1) : normalized;
  if (trimmed === '') return [];
  return trimmed.split('\n').map((line) => line.split('\t'));
}

export function toTsv(rows: string[][]): string {
  return rows.map((row) => row.map(escapeTsv).join('\t')).join('\n');
}

function escapeTsv(value: string): string {
  return value.replace(/\t/g, ' ').replace(/\r?\n/g, ' ');
}

export function toHtmlTable(rows: string[][]): string {
  const body = rows
    .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join('')}</tr>`)
    .join('');
  return `<table>${body}</table>`;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

export function parseHtmlTable(html: string): string[][] | null {
  if (!html || !/<\s*table/i.test(html)) return null;
  if (typeof DOMParser === 'undefined') return parseHtmlTableLoose(html);
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return null;
  const rows: string[][] = [];
  for (const tr of table.querySelectorAll('tr')) {
    const cells = [...tr.querySelectorAll('td,th')].map((cell) => cell.textContent?.replace(/\u00a0/g, ' ') ?? '');
    if (cells.length) rows.push(cells);
  }
  return rows.length ? rows : null;
}

function parseHtmlTableLoose(html: string): string[][] | null {
  const rows: string[][] = [];
  const rowRe = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi;
  let rowMatch: RegExpExecArray | null;
  while ((rowMatch = rowRe.exec(html))) {
    const cells: string[] = [];
    const cellRe = /<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/gi;
    let cellMatch: RegExpExecArray | null;
    while ((cellMatch = cellRe.exec(rowMatch[1] ?? ''))) {
      cells.push(decode(cellMatch[1]?.replace(/<[^>]+>/g, '') ?? ''));
    }
    if (cells.length) rows.push(cells);
  }
  return rows.length ? rows : null;
}

function decode(value: string): string {
  return value
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .trim();
}

export function parseClipboard(html: string, text: string): string[][] {
  const fromHtml = html ? parseHtmlTable(html) : null;
  if (fromHtml && fromHtml.length) return fromHtml;
  return parseTsv(text);
}

export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;
  const src = text.replace(/^\uFEFF/, '');
  for (let i = 0; i < src.length; i += 1) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"') {
        if (src[i + 1] === '"') {
          cell += '"';
          i += 1;
        } else {
          quoted = false;
        }
      } else {
        cell += ch;
      }
      continue;
    }
    if (ch === '"') {
      quoted = true;
      continue;
    }
    if (ch === ',') {
      row.push(cell);
      cell = '';
      continue;
    }
    if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && src[i + 1] === '\n') i += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
      continue;
    }
    cell += ch ?? '';
  }
  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }
  if (rows.length && rows[rows.length - 1]?.every((value) => value === '')) rows.pop();
  return rows;
}

export function toCsv(rows: string[][]): string {
  return rows.map((row) => row.map(escapeCsv).join(',')).join('\n');
}

function escapeCsv(value: string): string {
  if (/[",\n\r]/.test(value)) return `"${value.replace(/"/g, '""')}"`;
  return value;
}
