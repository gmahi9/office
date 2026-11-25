import { useEffect, useMemo, useRef, useState, type ClipboardEvent, type KeyboardEvent } from 'react';
import { formatA1, formatRangeA1, parseA1Range } from '../../domain/a1';
import { parseClipboard, toHtmlTable, toTsv } from '../../domain/clipboard';
import { computeFill, fillDownRange, fillRightRange } from '../../domain/fill';
import { offsetFormula } from '../../domain/formula';
import { formatValue } from '../../domain/format';
import { jumpCell } from '../../domain/navigate';
import {
  COL_COUNT,
  ROW_COUNT,
  cloneData,
  keyOf,
  normalizeRange,
  parseKey,
  rangeContains,
  usedBounds,
  type CellRange,
  type CellStyle,
  type ChartType,
  type CondType,
} from '../../domain/types';
import { saveLinkedRange } from '../../data/linkedRanges';
import type { UsageStat } from '../../data/usage';
import { COMMANDS, TOOLBAR_IDS, commandById } from '../../nav/commands';
import { rankCommands } from '../../nav/rank';
import type { WorkbookSession } from '../../engine/session';
import { toCsv } from '../../domain/clipboard';
import { sheetToDisplayRows, workbookFromCsv, workbookFromXlsx, workbookToXlsx } from '../../io/xlsx';
import { Dialog, ProgressButton } from '../../chrome/ui';
import { Grid, type GridHandle } from './Grid';

const SWATCHES = ['#1b3148', '#8d4520', '#2f7eb8', '#1f6b4a', '#8d2f2f', '#fffaf6', '#e7c3a8', '#d7ebf8'];

interface Clip {
  range: CellRange;
  rows: Array<Array<{ input?: string; style?: CellStyle }>>;
  text: string;
  cut: boolean;
}

export function SheetRoom({
  session,
  usage,
  reducedMotion,
  onUse,
  onDirty,
  onToast,
  onImported,
}: {
  session: WorkbookSession;
  usage: UsageStat[];
  reducedMotion: boolean;
  onUse: (id: string) => void;
  onDirty: () => void;
  onToast: (message: string, undo?: () => void) => void;
  onImported: (models: ReturnType<typeof workbookFromCsv>[]) => void;
}) {
  const sheet = session.activeSheet();
  const gridRef = useRef<GridHandle>(null);
  const clipRef = useRef<Clip | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const [rev, setRev] = useState(0);
  const [active, setActive] = useState({ r: 0, c: 0 });
  const [anchor, setAnchor] = useState({ r: 0, c: 0 });
  const [selection, setSelection] = useState<CellRange>({ r1: 0, c1: 0, r2: 0, c2: 0 });
  const [editing, setEditing] = useState<{ r: number; c: number; draft: string; focus?: boolean } | null>(null);
  const editingRef = useRef(editing);
  const draftRef = useRef('');
  editingRef.current = editing;
  const [trace, setTrace] = useState<Array<{ r: number; c: number }>>([]);
  const [traceToken, setTraceToken] = useState(0);
  const [sweep, setSweep] = useState<CellRange | null>(null);
  const [sweepToken, setSweepToken] = useState(0);
  const [dialog, setDialog] = useState<string | null>(null);
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null);
  const [filterCol, setFilterCol] = useState<{ col: number; x: number; y: number } | null>(null);
  const [colorTarget, setColorTarget] = useState<'fill' | 'font' | null>(null);
  const [fillColor, setFillColor] = useState('#e7c3a8');
  const [fontColor, setFontColor] = useState('#1b3148');
  const [shredId, setShredId] = useState<string | null>(null);
  const [query, setQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [hits, setHits] = useState<Array<{ sheetId: string; r: number; c: number }>>([]);
  const [hitAt, setHitAt] = useState(0);
  const [chartType, setChartType] = useState<ChartType>('bar');
  const [ruleType, setRuleType] = useState<CondType>('gt');
  const [ruleValue, setRuleValue] = useState('0');
  const [ruleValue2, setRuleValue2] = useState('10');
  const [listText, setListText] = useState('Technology, Healthcare, Finance');
  const [paletteOpen, setPaletteOpen] = useState(false);
  const [paletteQuery, setPaletteQuery] = useState('');

  function bump() {
    setRev((value) => value + 1);
    onDirty();
  }

  useEffect(() => {
    gridRef.current?.focus();
  }, [sheet.id]);

  const tools = useMemo(() => {
    const defs = TOOLBAR_IDS.map((id) => commandById(id)).filter((item): item is NonNullable<typeof item> => Boolean(item));
    return rankCommands(defs, usage, Date.now());
  }, [usage]);

  function selectCell(r: number, c: number, extend: boolean) {
    const resolved = session.anchor(sheet.id, r, c);
    if (extend) {
      setSelection(normalizeRange({ r1: anchor.r, c1: anchor.c, r2: resolved.r, c2: resolved.c }));
      setActive({ r: resolved.r, c: resolved.c });
    } else {
      setAnchor({ r: resolved.r, c: resolved.c });
      setActive({ r: resolved.r, c: resolved.c });
      setSelection(resolved.range);
    }
    setMenu(null);
  }

  function moveTo(r: number, c: number, extend: boolean) {
    const next = { r: Math.max(0, Math.min(ROW_COUNT - 1, r)), c: Math.max(0, Math.min(COL_COUNT - 1, c)) };
    selectCell(next.r, next.c, extend);
    gridRef.current?.scrollTo(next.r, next.c);
  }

  function commit(move: 'down' | 'up' | 'left' | 'right' | 'stay') {
    const current = editingRef.current;
    if (!current) return;
    const draft = draftRef.current || current.draft;
    editingRef.current = null;
    draftRef.current = '';
    setEditing(null);
    const formula = draft.trim().startsWith('=');
    session.setCell(sheet.id, current.r, current.c, draft);
    if (formula) {
      const prec = session.precedents(sheet.id, current.r, current.c);
      setTrace(prec.flatMap((range) => {
        const cells = [];
        for (let r = range.r1; r <= Math.min(range.r2, range.r1 + 30); r += 1) {
          for (let c = range.c1; c <= Math.min(range.c2, range.c1 + 12); c += 1) cells.push({ r, c });
        }
        return cells;
      }));
      setTraceToken(Date.now());
    }
    if (move === 'down') moveTo(current.r + 1, current.c, false);
    if (move === 'up') moveTo(current.r - 1, current.c, false);
    if (move === 'left') moveTo(current.r, current.c - 1, false);
    if (move === 'right') moveTo(current.r, current.c + 1, false);
    bump();
    gridRef.current?.focus();
  }

  function runFill(target: CellRange) {
    const edits = computeFill(sheet, selection, target);
    session.applyRecords(sheet.id, edits.map((edit) => ({ r: edit.r, c: edit.c, record: edit.record })));
    setSelection(normalizeRange(target));
    setSweep(normalizeRange(target));
    setSweepToken(Date.now());
    onUse('fill-down');
    bump();
  }

  function applyFormat(patch: Partial<CellStyle>) {
    session.patchStyle(sheet.id, selection, patch);
    bump();
  }

  function toggleFlag(flag: 'bold' | 'italic' | 'underline' | 'wrap') {
    const current = sheet.cells[keyOf(active.r, active.c)]?.style?.[flag] === true;
    applyFormat({ [flag]: !current });
  }

  function command(id: string) {
    onUse(id);
    const area = normalizeRange(selection);
    if (id === 'undo') {
      session.undo();
      bump();
      return;
    }
    if (id === 'redo') {
      session.redo();
      bump();
      return;
    }
    if (id === 'bold' || id === 'italic' || id === 'underline') {
      toggleFlag(id);
      return;
    }
    if (id === 'wrap') {
      toggleFlag('wrap');
      return;
    }
    if (id === 'align-left' || id === 'align-center' || id === 'align-right') {
      applyFormat({ align: id.slice(6) as 'left' | 'center' | 'right' });
      return;
    }
    if (id.startsWith('format-')) {
      applyFormat({ format: id.slice(7) });
      return;
    }
    if (id === 'fill-color') {
      setColorTarget('fill');
      return;
    }
    if (id === 'font-color') {
      setColorTarget('font');
      return;
    }
    if (id === 'border-all' || id === 'border-outer' || id === 'border-bottom' || id === 'border-none') {
      session.applyBorder(sheet.id, area, id.slice(7) as 'all' | 'outer' | 'bottom' | 'none');
      bump();
      return;
    }
    if (id === 'merge') {
      session.merge(sheet.id, area);
      bump();
      return;
    }
    if (id === 'freeze') {
      const frozen = sheet.freeze.rows === active.r && sheet.freeze.cols === active.c;
      session.setFreeze(sheet.id, frozen ? 0 : active.r, frozen ? 0 : active.c);
      bump();
      return;
    }
    if (id === 'sort-asc' || id === 'sort-desc') {
      session.sort(sheet.id, area, active.c, id === 'sort-asc' ? 'asc' : 'desc');
      bump();
      return;
    }
    if (id === 'filter') {
      session.toggleFilter(sheet.id, area.r1 === area.r2 && area.c1 === area.c2 ? usedBounds(sheet) : area);
      bump();
      return;
    }
    if (id === 'chart') {
      setDialog('chart');
      return;
    }
    if (id === 'conditional') {
      setDialog('rule');
      return;
    }
    if (id === 'validation') {
      setDialog('list');
      return;
    }
    if (id === 'name-range') {
      setDialog('name');
      return;
    }
    if (id === 'link-range') {
      void publishLink();
      return;
    }
    if (id === 'find') {
      setDialog('find');
      return;
    }
    if (id === 'replace') {
      setDialog('find');
      return;
    }
    if (id === 'insert-row') {
      session.insertRows(sheet.id, active.r, 1);
      bump();
      return;
    }
    if (id === 'insert-col') {
      session.insertCols(sheet.id, active.c, 1);
      bump();
      return;
    }
    if (id === 'delete-row') {
      session.deleteRows(sheet.id, area.r1, area.r2 - area.r1 + 1);
      bump();
      return;
    }
    if (id === 'delete-col') {
      session.deleteCols(sheet.id, area.c1, area.c2 - area.c1 + 1);
      bump();
      return;
    }
    if (id === 'hide-row') {
      session.hideRows(sheet.id, rangeIndexes(area.r1, area.r2));
      bump();
      return;
    }
    if (id === 'hide-col') {
      session.hideCols(sheet.id, rangeIndexes(area.c1, area.c2));
      bump();
      return;
    }
    if (id === 'unhide') {
      session.unhideAll(sheet.id);
      bump();
      return;
    }
    if (id === 'fill-down') runFill(fillDownRange(area));
    if (id === 'fill-right') runFill(fillRightRange(area));
  }

  async function publishLink() {
    const area = normalizeRange(selection);
    const id = crypto.randomUUID();
    await saveLinkedRange({
      id,
      workbookId: session.model.id,
      sheetId: sheet.id,
      sheetName: sheet.name,
      a1: formatRangeA1(area),
      r1: area.r1,
      c1: area.c1,
      r2: area.r2,
      c2: area.c2,
      cells: session.snapshotRange(sheet.id, area),
      createdAt: Date.now(),
      updatedAt: Date.now(),
    });
    onToast(`Linked ${formatRangeA1(area)} on this device.`);
  }

  function onKeyDown(event: KeyboardEvent<HTMLDivElement>) {
    const tag = (event.target as HTMLElement).tagName;
    if (tag === 'TEXTAREA' || tag === 'INPUT') return;
    if (editingRef.current) return;
    const mod = event.ctrlKey || event.metaKey;
    const key = event.key;
    if (mod && key.toLowerCase() === 'b') {
      event.preventDefault();
      command('bold');
      return;
    }
    if (mod && key.toLowerCase() === 'i') {
      event.preventDefault();
      command('italic');
      return;
    }
    if (mod && key.toLowerCase() === 'u') {
      event.preventDefault();
      command('underline');
      return;
    }
    if (mod && key.toLowerCase() === 'z') {
      event.preventDefault();
      command(event.shiftKey ? 'redo' : 'undo');
      return;
    }
    if (mod && key.toLowerCase() === 'y') {
      event.preventDefault();
      command('redo');
      return;
    }
    if (mod && key.toLowerCase() === 'f') {
      event.preventDefault();
      command('find');
      return;
    }
    if (mod && key.toLowerCase() === 'h') {
      event.preventDefault();
      command('replace');
      return;
    }
    if (mod && key.toLowerCase() === 'k') {
      event.preventDefault();
      setPaletteOpen(true);
      return;
    }
    if (mod && key.toLowerCase() === 'd') {
      event.preventDefault();
      command('fill-down');
      return;
    }
    if (mod && key.toLowerCase() === 'r') {
      event.preventDefault();
      command('fill-right');
      return;
    }
    if (mod && key.toLowerCase() === 'a') {
      event.preventDefault();
      const used = usedBounds(sheet);
      const already = selection.r1 === used.r1 && selection.c1 === used.c1 && selection.r2 === used.r2 && selection.c2 === used.c2;
      setSelection(already ? { r1: 0, c1: 0, r2: ROW_COUNT - 1, c2: COL_COUNT - 1 } : used);
      return;
    }
    if (key === 'F2') {
      event.preventDefault();
      draftRef.current = session.inputAt(sheet.id, active.r, active.c);
      setEditing({ r: active.r, c: active.c, draft: draftRef.current, focus: true });
      return;
    }
    if (key === 'Delete' || key === 'Backspace') {
      event.preventDefault();
      session.clear(sheet.id, selection, 'values');
      bump();
      return;
    }
    if (key === 'Enter') {
      event.preventDefault();
      moveTo(active.r + (event.shiftKey ? -1 : 1), active.c, false);
      return;
    }
    if (key === 'Tab') {
      event.preventDefault();
      moveTo(active.r, active.c + (event.shiftKey ? -1 : 1), false);
      return;
    }
    if (key === 'Home') {
      event.preventDefault();
      moveTo(mod ? 0 : active.r, 0, event.shiftKey);
      return;
    }
    if (key === 'End') {
      event.preventDefault();
      const used = usedBounds(sheet);
      moveTo(mod ? used.r2 : active.r, mod ? used.c2 : used.c2, event.shiftKey);
      return;
    }
    if (key === 'PageDown' || key === 'PageUp') {
      event.preventDefault();
      moveTo(active.r + (key === 'PageDown' ? 25 : -25), active.c, event.shiftKey);
      return;
    }
    const arrow = arrowDelta(key);
    if (arrow) {
      event.preventDefault();
      if (mod) {
        const jumped = jumpCell(
          (r, c) => session.isEmpty(sheet.id, r, c),
          (r, c) => sheet.hiddenRows.includes(r) || sheet.filterHidden.includes(r) || sheet.hiddenCols.includes(c),
          active.r,
          active.c,
          arrow.dr,
          arrow.dc,
        );
        moveTo(jumped.r, jumped.c, event.shiftKey);
      } else {
        moveTo(active.r + arrow.dr, active.c + arrow.dc, event.shiftKey);
      }
      return;
    }
    if (key.length === 1 && !mod && !event.altKey) {
      event.preventDefault();
      draftRef.current = key;
      setEditing({ r: active.r, c: active.c, draft: key, focus: true });
    }
  }

  function buildGrid(area: CellRange, inputs: boolean): string[][] {
    const rows: string[][] = [];
    for (let r = area.r1; r <= area.r2; r += 1) {
      const row: string[] = [];
      for (let c = area.c1; c <= area.c2; c += 1) {
        row.push(inputs ? session.inputAt(sheet.id, r, c) || session.display(sheet.id, r, c) : session.display(sheet.id, r, c));
      }
      rows.push(row);
    }
    return rows;
  }

  function onCopy(event: ClipboardEvent<HTMLDivElement>, cut = false) {
    event.preventDefault();
    const area = normalizeRange(selection);
    const rows = buildGrid(area, true);
    const rich = rows.map((row, r) => row.map((_, c) => {
      const cell = sheet.cells[keyOf(area.r1 + r, area.c1 + c)];
      return { input: cell?.input, style: cell?.style };
    }));
    const text = toTsv(rows);
    clipRef.current = { range: area, rows: rich, text, cut };
    event.clipboardData.setData('text/plain', text);
    event.clipboardData.setData('text/html', toHtmlTable(buildGrid(area, false)));
    if (cut) {
      session.clear(sheet.id, area, 'all');
      bump();
    }
  }

  function onPaste(event: ClipboardEvent<HTMLDivElement>) {
    event.preventDefault();
    const text = event.clipboardData.getData('text/plain');
    const html = event.clipboardData.getData('text/html');
    const clip = clipRef.current;
    if (clip && text === clip.text) {
      const edits = clip.rows.flatMap((row, r) => row.map((cell, c) => ({
        r: active.r + r,
        c: active.c + c,
        record: cell.input || cell.style
          ? {
              input: cell.input?.startsWith('=')
                ? offsetImported(cell.input, active.r - clip.range.r1, active.c - clip.range.c1)
                : cell.input,
              style: cell.style,
            }
          : null,
      })));
      session.applyRecords(sheet.id, edits);
      if (clip.cut) clipRef.current = null;
    } else {
      const grid = parseClipboard(html, text);
      const edits = grid.flatMap((row, r) => row.map((value, c) => ({
        r: active.r + r,
        c: active.c + c,
        record: value === '' ? null : { input: value },
      })));
      session.applyRecords(sheet.id, edits);
    }
    bump();
  }

  function search() {
    const needle = matchCase ? query : query.toLowerCase();
    if (!needle) {
      setHits([]);
      return;
    }
    const found: Array<{ sheetId: string; r: number; c: number }> = [];
    for (const item of session.model.sheets) {
      for (const [key, cell] of Object.entries(item.cells)) {
        const pos = parseKey(key);
        const sources = [cell.input ?? '', session.display(item.id, pos.r, pos.c)];
        const matched = sources.some((value) => (matchCase ? value : value.toLowerCase()).includes(needle));
        if (matched) found.push({ sheetId: item.id, r: pos.r, c: pos.c });
      }
    }
    setHits(found);
    setHitAt(0);
    const first = found[0];
    if (first) reveal(first);
  }

  function reveal(hit: { sheetId: string; r: number; c: number }) {
    if (hit.sheetId !== sheet.id) {
      session.setActive(hit.sheetId);
      bump();
    }
    setActive({ r: hit.r, c: hit.c });
    setAnchor({ r: hit.r, c: hit.c });
    setSelection({ r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c });
    gridRef.current?.scrollTo(hit.r, hit.c);
  }

  function replaceOne() {
    const hit = hits[hitAt];
    if (!hit) return;
    const input = session.inputAt(hit.sheetId, hit.r, hit.c);
    session.setCell(hit.sheetId, hit.r, hit.c, replaceText(input, query, replacement, matchCase));
    bump();
    search();
  }

  function replaceAll() {
    if (!query) return;
    let count = 0;
    session.beginGesture();
    for (const item of session.model.sheets) {
      for (const [key, cell] of Object.entries(item.cells)) {
        if (!cell.input || !includesText(cell.input, query, matchCase)) continue;
        const pos = parseKey(key);
        session.setCell(item.id, pos.r, pos.c, replaceText(cell.input, query, replacement, matchCase));
        count += 1;
      }
    }
    session.endGesture();
    bump();
    onToast(count ? `Replaced ${count}.` : 'Nothing matched.');
    search();
  }

  async function saveXlsx(report: (value: number) => void, picker: boolean) {
    report(0.2);
    const bytes = await workbookToXlsx(session.model);
    report(0.72);
    await saveBytes(`${session.model.title}.xlsx`, bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', picker);
    report(1);
  }

  async function exportCsv(report: (value: number) => void) {
    report(0.35);
    const rows = sheetToDisplayRows(sheet, (r, c) => session.display(sheet.id, r, c));
    const text = toCsv(rows);
    report(0.8);
    await saveBytes(`${session.model.title}-${sheet.name}.csv`, new TextEncoder().encode(text), 'text/csv', false);
    report(1);
  }

  async function openFile(file: File) {
    const lower = file.name.toLowerCase();
    if (lower.endsWith('.csv')) {
      onImported([workbookFromCsv(file.name, await file.text())]);
      return;
    }
    const model = await workbookFromXlsx(await file.arrayBuffer(), file.name.replace(/\.xlsx$/i, ''));
    onImported([model]);
  }

  function removeSheet(id: string) {
    if (session.model.sheets.length <= 1) {
      onToast('A workbook keeps at least one sheet.');
      return;
    }
    const snapshot = cloneData(session.model);
    const name = session.sheet(id)?.name ?? 'Sheet';
    setShredId(id);
    window.setTimeout(() => {
      session.deleteSheet(id);
      setShredId(null);
      setActive({ r: 0, c: 0 });
      setSelection({ r1: 0, c1: 0, r2: 0, c2: 0 });
      bump();
      onToast(`${name} was shredded.`, () => {
        session.replaceModel(snapshot);
        bump();
      });
    }, reducedMotion ? 0 : 340);
  }

  const area = normalizeRange(selection);
  let count = 0;
  let numeric = 0;
  let sum = 0;
  for (const key of Object.keys(sheet.cells)) {
    const pos = parseKey(key);
    if (!rangeContains(area, pos.r, pos.c)) continue;
    const value = session.raw(sheet.id, pos.r, pos.c);
    if (value == null || value === '') continue;
    count += 1;
    if (typeof value === 'number') {
      numeric += 1;
      sum += value;
    }
  }
  const name = session.nameForRange(sheet.id, area);
  const rankedPalette = rankCommands(
    COMMANDS.filter((item) => !paletteQuery || item.label.toLowerCase().includes(paletteQuery.toLowerCase())),
    usage,
    Date.now(),
  );

  return (
    <section className="room" aria-label={session.model.title}>
      <div className="toolbar" role="toolbar" aria-label="Sheet tools">
        <button type="button" className="tool" onClick={() => command('undo')} disabled={!session.canUndo} aria-label="Undo">Undo</button>
        <button type="button" className="tool" onClick={() => command('redo')} disabled={!session.canRedo} aria-label="Redo">Redo</button>
        <span className="tool-gap" />
        <ProgressButton label="Save" doneLabel="Saved" onRun={(report) => saveXlsx(report, true)} />
        <ProgressButton label="Export" doneLabel="Exported" onRun={exportCsv} />
        <ProgressButton label="Download" doneLabel="Downloaded" onRun={(report) => saveXlsx(report, false)} />
        <button type="button" className="tool" onClick={() => fileRef.current?.click()}>Open</button>
        <input
          ref={fileRef}
          type="file"
          accept=".xlsx,.csv,text/csv"
          hidden
          onChange={(event) => {
            const file = event.target.files?.[0];
            event.target.value = '';
            if (file) void openFile(file);
          }}
        />
        <span className="tool-gap" />
        <select
          aria-label="Font"
          value={sheet.cells[keyOf(active.r, active.c)]?.style?.fontFamily ?? 'Outfit'}
          onChange={(event) => applyFormat({ fontFamily: event.target.value })}
        >
          <option>Outfit</option>
          <option>Syne</option>
          <option>mono</option>
        </select>
        <select
          aria-label="Font size"
          value={sheet.cells[keyOf(active.r, active.c)]?.style?.fontSize ?? 13}
          onChange={(event) => applyFormat({ fontSize: Number(event.target.value) })}
        >
          {[11, 12, 13, 14, 16].map((size) => <option key={size}>{size}</option>)}
        </select>
        {tools.map((tool) => (
          <button key={tool.id} type="button" className="tool" onClick={() => command(tool.id)} title={tool.shortcut}>
            {shortLabel(tool.id)}
          </button>
        ))}
        <button type="button" className="tool" onClick={() => setPaletteOpen(true)}>Commands</button>
      </div>
      <form
        className="formula-bar"
        onSubmit={(event) => {
          event.preventDefault();
          if (editing) commit('stay');
        }}
      >
        <input
          className="name-box"
          aria-label="Name box"
          key={`${active.r},${active.c},${name ?? ''},${rev}`}
          defaultValue={name ?? formatA1(active.r, active.c)}
          onBlur={(event) => {
            const text = event.target.value.trim();
            const parsed = parseA1Range(text);
            if (parsed) {
              setActive({ r: parsed.r1, c: parsed.c1 });
              setAnchor({ r: parsed.r1, c: parsed.c1 });
              setSelection(parsed);
              gridRef.current?.scrollTo(parsed.r1, parsed.c1);
              return;
            }
            if (text && session.defineName(text, sheet.id, selection)) {
              onUse('name-range');
              bump();
            }
          }}
        />
        <label className="sr-only" htmlFor="formula-input">Formula</label>
        <input
          id="formula-input"
          className="formula-input"
          value={editing ? editing.draft : session.inputAt(sheet.id, active.r, active.c)}
          onFocus={() => {
            if (!editingRef.current) {
              draftRef.current = session.inputAt(sheet.id, active.r, active.c);
              setEditing({ r: active.r, c: active.c, draft: draftRef.current, focus: false });
            }
          }}
          onChange={(event) => {
            draftRef.current = event.target.value;
            setEditing({ r: active.r, c: active.c, draft: event.target.value, focus: false });
          }}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              commit('down');
            } else if (event.key === 'Escape') {
              editingRef.current = null;
              setEditing(null);
              gridRef.current?.focus();
            }
          }}
        />
      </form>
      <Grid
        ref={gridRef}
        session={session}
        rev={rev}
        sheetId={sheet.id}
        selection={selection}
        active={active}
        editing={editing}
        trace={trace}
        traceToken={traceToken}
        sweep={sweep}
        sweepToken={sweepToken}
        findHit={hits[hitAt] && hits[hitAt].sheetId === sheet.id ? hits[hitAt] : null}
        reducedMotion={reducedMotion}
        onSelect={selectCell}
        onKeyDown={onKeyDown}
        onEditStart={(r, c, seed) => {
          draftRef.current = seed ?? session.inputAt(sheet.id, r, c);
          setEditing({ r, c, draft: draftRef.current, focus: true });
        }}
        onEditDraft={(draft) => {
          draftRef.current = draft;
          setEditing((current) => (current ? { ...current, draft } : current));
        }}
        onCommit={commit}
        onCancel={() => {
          editingRef.current = null;
          setEditing(null);
        }}
        onFill={runFill}
        onCopy={(event) => onCopy(event, false)}
        onCut={(event) => onCopy(event, true)}
        onPaste={onPaste}
        onResizeCol={(col, width) => {
          session.setColWidth(sheet.id, col, width);
          setRev((value) => value + 1);
        }}
        onResizeRow={(row, height) => {
          session.setRowHeight(sheet.id, row, height);
          setRev((value) => value + 1);
        }}
        onGestureStart={() => session.beginGesture()}
        onGestureEnd={() => {
          session.endGesture();
          bump();
        }}
        onContext={(r, c, x, y) => {
          selectCell(r, c, false);
          setMenu({ x, y });
        }}
        onFilter={(col, x, y) => setFilterCol({ col, x, y })}
        onChartMove={(id, x, y) => {
          session.updateChart(sheet.id, id, { x, y });
          setRev((value) => value + 1);
        }}
        onChartRemove={(id) => {
          session.removeChart(sheet.id, id);
          bump();
        }}
      />
      <div className="sheet-tabs" role="tablist" aria-label="Sheets">
        {session.model.sheets.map((item) => (
          <div
            key={item.id}
            role="tab"
            tabIndex={0}
            aria-selected={item.id === sheet.id}
            className={`sheet-tab${item.id === sheet.id ? ' is-active' : ''}${shredId === item.id ? ' is-shred' : ''}`}
            onClick={() => {
              session.setActive(item.id);
              setActive({ r: 0, c: 0 });
              setSelection({ r1: 0, c1: 0, r2: 0, c2: 0 });
              bump();
            }}
            onKeyDown={(event) => {
              if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault();
                session.setActive(item.id);
                bump();
              }
            }}
            onDoubleClick={() => {
              const next = window.prompt('Sheet name', item.name);
              if (next && session.renameSheet(item.id, next)) bump();
            }}
          >
            {item.name}
            <button
              type="button"
              className="tab-x"
              aria-label={`Delete ${item.name}`}
              onClick={(event) => {
                event.stopPropagation();
                removeSheet(item.id);
              }}
            >
              ×
            </button>
          </div>
        ))}
        <button type="button" className="tool" onClick={() => { session.addSheet(); bump(); }} aria-label="Add sheet">+</button>
      </div>
      <footer className="status">
        <span>{session.dirty ? 'Saving on this device…' : 'Saved on this device'}</span>
        <span>{count} selected{numeric ? ` · sum ${formatValue(sum, 'number')} · avg ${formatValue(sum / numeric, 'number')}` : ''}</span>
        {sheet.filter ? <button type="button" className="text-btn" onClick={() => command('filter')}>Clear filter</button> : null}
      </footer>
      <p className="sr-only" aria-live="polite">
        {formatA1(active.r, active.c)} {session.display(sheet.id, active.r, active.c)}
      </p>
      {menu ? (
        <div className="menu" style={{ left: menu.x, top: menu.y }} role="menu">
          {['insert-row', 'delete-row', 'insert-col', 'delete-col', 'hide-row', 'hide-col', 'unhide'].map((id) => (
            <button key={id} type="button" role="menuitem" onClick={() => { command(id); setMenu(null); }}>{commandById(id)?.label}</button>
          ))}
        </div>
      ) : null}
      {filterCol && sheet.filter ? (
        <FilterPopover
          values={session.distinctValues(sheet.id, filterCol.col, sheet.filter.range)}
          allowed={sheet.filter.criteria[filterCol.col] ?? null}
          x={filterCol.x}
          y={filterCol.y}
          onClose={() => setFilterCol(null)}
          onApply={(allowed) => {
            session.setFilterValues(sheet.id, filterCol.col, allowed);
            setFilterCol(null);
            bump();
          }}
        />
      ) : null}
      {colorTarget ? (
        <div className="popover swatches" role="dialog" aria-label={colorTarget === 'fill' ? 'Fill color' : 'Text color'}>
          {SWATCHES.map((color) => (
            <button
              key={color}
              type="button"
              style={{ background: color }}
              aria-label={color}
              onClick={() => {
                if (colorTarget === 'fill') {
                  setFillColor(color);
                  applyFormat({ fill: color });
                  onUse('fill-color');
                } else {
                  setFontColor(color);
                  applyFormat({ color });
                  onUse('font-color');
                }
                setColorTarget(null);
              }}
            />
          ))}
          <input
            type="color"
            aria-label="Custom color"
            value={colorTarget === 'fill' ? fillColor : fontColor}
            onChange={(event) => {
              const color = event.target.value;
              if (colorTarget === 'fill') {
                setFillColor(color);
                applyFormat({ fill: color });
              } else {
                setFontColor(color);
                applyFormat({ color });
              }
            }}
          />
          <button type="button" className="text-btn" onClick={() => setColorTarget(null)}>Done</button>
        </div>
      ) : null}
      {dialog === 'find' ? (
        <Dialog title="Find and replace" onClose={() => setDialog(null)}>
          <label>Find<input value={query} onChange={(event) => setQuery(event.target.value)} autoFocus /></label>
          <label>Replace<input value={replacement} onChange={(event) => setReplacement(event.target.value)} /></label>
          <label className="check"><input type="checkbox" checked={matchCase} onChange={(event) => setMatchCase(event.target.checked)} /> Match case</label>
          <div className="dialog-actions">
            <button type="button" className="tool" onClick={search}>Find</button>
            <button type="button" className="tool" onClick={() => { const next = hits[(hitAt + 1) % Math.max(1, hits.length)]; if (next) { setHitAt((hitAt + 1) % hits.length); reveal(next); } }}>Next</button>
            <button type="button" className="tool" onClick={replaceOne}>Replace</button>
            <button type="button" className="tool" onClick={replaceAll}>Replace all</button>
          </div>
          <p>{hits.length ? `${hitAt + 1} of ${hits.length}` : query ? 'No matches yet.' : 'Search formulas and values.'}</p>
        </Dialog>
      ) : null}
      {dialog === 'chart' ? (
        <Dialog title="Chart from selection" onClose={() => setDialog(null)}>
          <label>Type
            <select value={chartType} onChange={(event) => setChartType(event.target.value as ChartType)}>
              <option value="bar">Bar</option>
              <option value="line">Line</option>
              <option value="pie">Pie</option>
              <option value="scatter">Scatter</option>
            </select>
          </label>
          <div className="dialog-actions">
            <button type="button" className="tool" onClick={() => {
              session.addChart(sheet.id, chartType, selection, `${chartType[0].toUpperCase()}${chartType.slice(1)} chart`);
              setDialog(null);
              bump();
            }}>Insert</button>
          </div>
        </Dialog>
      ) : null}
      {dialog === 'rule' ? (
        <Dialog title="Conditional format" onClose={() => setDialog(null)}>
          <label>Rule
            <select value={ruleType} onChange={(event) => setRuleType(event.target.value as CondType)}>
              <option value="gt">Greater than</option>
              <option value="lt">Less than</option>
              <option value="eq">Equal to</option>
              <option value="between">Between</option>
              <option value="contains">Text contains</option>
              <option value="colorScale">Color scale</option>
            </select>
          </label>
          {ruleType !== 'colorScale' ? <label>Value<input value={ruleValue} onChange={(event) => setRuleValue(event.target.value)} /></label> : null}
          {ruleType === 'between' ? <label>And<input value={ruleValue2} onChange={(event) => setRuleValue2(event.target.value)} /></label> : null}
          <div className="dialog-actions">
            <button type="button" className="tool" onClick={() => {
              session.addConditional(sheet.id, {
                id: crypto.randomUUID(),
                range: normalizeRange(selection),
                type: ruleType,
                value: ruleValue,
                value2: ruleValue2,
                fill: ruleType === 'colorScale' ? undefined : '#f3e1d4',
                color: ruleType === 'colorScale' ? undefined : '#8d4520',
                minColor: '#d7ebf8',
                maxColor: '#8d4520',
              });
              setDialog(null);
              bump();
            }}>Apply</button>
          </div>
          <ul className="plain-list">
            {sheet.conditional.map((rule) => (
              <li key={rule.id}>
                {rule.type} {rule.value ?? ''}
                <button type="button" className="text-btn" onClick={() => { session.removeConditional(sheet.id, rule.id); bump(); }}>Remove</button>
              </li>
            ))}
          </ul>
        </Dialog>
      ) : null}
      {dialog === 'list' ? (
        <Dialog title="List validation" onClose={() => setDialog(null)}>
          <label>Allowed values<input value={listText} onChange={(event) => setListText(event.target.value)} /></label>
          <div className="dialog-actions">
            <button type="button" className="tool" onClick={() => {
              const values = listText.split(',').map((item) => item.trim()).filter(Boolean);
              session.addValidation(sheet.id, { id: crypto.randomUUID(), range: normalizeRange(selection), type: 'list', values });
              setDialog(null);
              bump();
            }}>Apply</button>
          </div>
        </Dialog>
      ) : null}
      {dialog === 'name' ? (
        <Dialog title="Named ranges" onClose={() => setDialog(null)}>
          <form onSubmit={(event) => {
            event.preventDefault();
            const data = new FormData(event.currentTarget);
            const next = String(data.get('name') ?? '');
            if (session.defineName(next, sheet.id, selection)) {
              bump();
              event.currentTarget.reset();
            }
          }}>
            <label>Name<input name="name" autoFocus /></label>
            <button type="submit" className="tool">Define {formatRangeA1(normalizeRange(selection))}</button>
          </form>
          <ul className="plain-list">
            {session.model.names.map((item) => (
              <li key={item.name}>
                {item.name}
                <button type="button" className="text-btn" onClick={() => { session.removeName(item.name); bump(); }}>Remove</button>
              </li>
            ))}
          </ul>
        </Dialog>
      ) : null}
      {paletteOpen ? (
        <Dialog title="Commands" onClose={() => setPaletteOpen(false)}>
          <input
            aria-label="Filter commands"
            value={paletteQuery}
            onChange={(event) => setPaletteQuery(event.target.value)}
            placeholder="Search"
            autoFocus
          />
          <ul className="command-list">
            {rankedPalette.map((item) => (
              <li key={item.id}>
                <button type="button" onClick={() => { command(item.id); setPaletteOpen(false); }}>
                  <span>{item.label}</span>
                  {item.shortcut ? <kbd>{item.shortcut}</kbd> : null}
                </button>
              </li>
            ))}
          </ul>
          <button type="button" className="text-btn" onClick={() => onUse('reset-usage')}>Reset toolbar order</button>
        </Dialog>
      ) : null}
      </section>
  );
}

function FilterPopover({
  values,
  allowed,
  x,
  y,
  onClose,
  onApply,
}: {
  values: string[];
  allowed: string[] | null;
  x: number;
  y: number;
  onClose: () => void;
  onApply: (allowed: string[] | null) => void;
}) {
  const [picked, setPicked] = useState<string[]>(allowed ?? values);
  return (
    <div className="popover filter-pop" style={{ left: x, top: y }} role="dialog" aria-label="Filter values">
      <div className="filter-list">
        {values.map((value) => (
          <label key={value || '(blank)'}>
            <input
              type="checkbox"
              checked={picked.includes(value)}
              onChange={(event) => {
                setPicked((current) => event.target.checked ? [...current, value] : current.filter((item) => item !== value));
              }}
            />
            {value || '(blank)'}
          </label>
        ))}
      </div>
      <div className="dialog-actions">
        <button type="button" className="tool" onClick={() => onApply(picked.length === values.length ? null : picked)}>Apply</button>
        <button type="button" className="text-btn" onClick={onClose}>Close</button>
      </div>
    </div>
  );
}

function shortLabel(id: string): string {
  const labels: Record<string, string> = {
    bold: 'Bold',
    italic: 'Italic',
    underline: 'Underline',
    'align-left': 'Left',
    'align-center': 'Center',
    'align-right': 'Right',
    wrap: 'Wrap',
    'format-currency': 'USD',
    'format-percent': 'Percent',
    'format-number': 'Number',
    'format-date': 'Date',
    'fill-color': 'Fill',
    'font-color': 'Text',
    'border-all': 'Borders',
    merge: 'Merge',
    freeze: 'Freeze',
    'sort-asc': 'Sort A–Z',
    'sort-desc': 'Sort Z–A',
    filter: 'Filter',
    chart: 'Chart',
    conditional: 'Rule',
    validation: 'List',
    'link-range': 'Link',
  };
  return labels[id] ?? id;
}

function arrowDelta(key: string): { dr: number; dc: number } | null {
  if (key === 'ArrowUp') return { dr: -1, dc: 0 };
  if (key === 'ArrowDown') return { dr: 1, dc: 0 };
  if (key === 'ArrowLeft') return { dr: 0, dc: -1 };
  if (key === 'ArrowRight') return { dr: 0, dc: 1 };
  return null;
}

function rangeIndexes(start: number, end: number): number[] {
  const list = [];
  for (let i = start; i <= end; i += 1) list.push(i);
  return list;
}

function includesText(source: string, query: string, matchCase: boolean): boolean {
  return (matchCase ? source : source.toLowerCase()).includes(matchCase ? query : query.toLowerCase());
}

function replaceText(source: string, query: string, replacement: string, matchCase: boolean): string {
  const flags = matchCase ? 'g' : 'gi';
  return source.replace(new RegExp(escapeRegExp(query), flags), replacement);
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function offsetImported(formula: string, dRow: number, dCol: number): string {
  return dRow || dCol ? offsetFormula(formula, dRow, dCol) : formula;
}

async function saveBytes(filename: string, bytes: Uint8Array, mime: string, picker: boolean) {
  const copy = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(copy).set(bytes);
  const file = new File([copy], filename, { type: mime });
  const host = window as Window & {
    showSaveFilePicker?: (options: { suggestedName?: string }) => Promise<{
      createWritable: () => Promise<{ write: (data: Blob) => Promise<void>; close: () => Promise<void> }>;
    }>;
  };
  if (picker && host.showSaveFilePicker) {
    try {
      const handle = await host.showSaveFilePicker({ suggestedName: filename });
      const writable = await handle.createWritable();
      await writable.write(file);
      await writable.close();
      return;
    } catch (error) {
      if ((error as { name?: string }).name === 'AbortError') return;
    }
  }
  const url = URL.createObjectURL(file);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}
