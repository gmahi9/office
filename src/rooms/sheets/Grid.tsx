import {
  forwardRef,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type KeyboardEvent as ReactKeyboardEvent,
  type PointerEvent as ReactPointerEvent,
  type WheelEvent as ReactWheelEvent,
  type ReactNode,
} from 'react';
import { colToA } from '../../domain/a1';
import {
  COL_COUNT,
  COL_HEADER_HEIGHT,
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  ROW_COUNT,
  ROW_HEADER_WIDTH,
  findMerge,
  keyOf,
  normalizeRange,
  rangeContains,
  type CellRange,
  type CellStyle,
} from '../../domain/types';
import { SizeIndex, visibleSpan } from '../../domain/sizeIndex';
import type { WorkbookSession } from '../../engine/session';
import { ChartCard } from './charts';

export interface GridHandle {
  scrollTo(r: number, c: number): void;
  focus(): void;
}

interface GridProps {
  session: WorkbookSession;
  rev: number;
  sheetId: string;
  selection: CellRange;
  active: { r: number; c: number };
  editing: { r: number; c: number; draft: string; focus?: boolean } | null;
  trace: Array<{ r: number; c: number }>;
  traceToken: number;
  sweep: CellRange | null;
  sweepToken: number;
  findHit: { r: number; c: number } | null;
  reducedMotion: boolean;
  onSelect: (r: number, c: number, extend: boolean) => void;
  onKeyDown: (event: ReactKeyboardEvent<HTMLDivElement>) => void;
  onEditStart: (r: number, c: number, seed?: string) => void;
  onEditDraft: (draft: string) => void;
  onCommit: (move: 'down' | 'up' | 'left' | 'right' | 'stay') => void;
  onCancel: () => void;
  onFill: (target: CellRange) => void;
  onCopy: (event: React.ClipboardEvent<HTMLDivElement>) => void;
  onCut: (event: React.ClipboardEvent<HTMLDivElement>) => void;
  onPaste: (event: React.ClipboardEvent<HTMLDivElement>) => void;
  onResizeCol: (col: number, width: number) => void;
  onResizeRow: (row: number, height: number) => void;
  onGestureStart: () => void;
  onGestureEnd: () => void;
  onContext: (r: number, c: number, x: number, y: number) => void;
  onFilter: (col: number, x: number, y: number) => void;
  onChartMove: (id: string, x: number, y: number) => void;
  onChartRemove: (id: string) => void;
}

export const Grid = forwardRef<GridHandle, GridProps>(function Grid(props, ref) {
  const rootRef = useRef<HTMLDivElement>(null);
  const bodyRef = useRef<HTMLDivElement>(null);
  const topRef = useRef<HTMLDivElement>(null);
  const leftRef = useRef<HTMLDivElement>(null);
  const colRef = useRef<HTMLDivElement>(null);
  const rowRef = useRef<HTMLDivElement>(null);
  const rowsRef = useRef<SizeIndex | null>(null);
  const colsRef = useRef<SizeIndex | null>(null);
  const drag = useRef<'select' | 'fill' | null>(null);
  const editorNode = useRef<HTMLTextAreaElement | null>(null);
  const [win, setWin] = useState({ top: 0, left: 0, width: 900, height: 600 });
  const [fillTo, setFillTo] = useState<CellRange | null>(null);
  const [lines, setLines] = useState<Array<{ x1: number; y1: number; x2: number; y2: number }>>([]);
  const sheet = props.session.sheet(props.sheetId) ?? props.session.activeSheet();
  const { rows, cols } = useMemo(() => {
    const hiddenRows = [...sheet.hiddenRows, ...sheet.filterHidden];
    return {
      rows: new SizeIndex(ROW_COUNT, DEFAULT_ROW_HEIGHT, hiddenRows, sheet.rowHeights),
      cols: new SizeIndex(COL_COUNT, DEFAULT_COL_WIDTH, sheet.hiddenCols, sheet.colWidths),
    };
  }, [props.rev, props.sheetId, sheet]);
  rowsRef.current = rows;
  colsRef.current = cols;
  const freezeR = sheet.freeze.rows;
  const freezeC = sheet.freeze.cols;
  const frozenH = rows.offset(freezeR);
  const frozenW = cols.offset(freezeC);
  const bounds = useMemo(() => props.session.scaleMap(props.sheetId), [props.rev, props.sheetId, props.session]);

  useImperativeHandle(ref, () => ({
    focus() {
      rootRef.current?.focus();
    },
    scrollTo(r: number, c: number) {
      const body = bodyRef.current;
      const rowIndex = rowsRef.current;
      const colIndex = colsRef.current;
      if (!body || !rowIndex || !colIndex) return;
      if (r >= freezeR) {
        const top = rowIndex.offset(r) - rowIndex.offset(freezeR);
        const bottom = top + rowIndex.size(r);
        if (top < body.scrollTop) body.scrollTop = top;
        else if (bottom > body.scrollTop + body.clientHeight) body.scrollTop = bottom - body.clientHeight + 8;
      }
      if (c >= freezeC) {
        const left = colIndex.offset(c) - colIndex.offset(freezeC);
        const right = left + colIndex.size(c);
        if (left < body.scrollLeft) body.scrollLeft = left;
        else if (right > body.scrollLeft + body.clientWidth) body.scrollLeft = right - body.clientWidth + 8;
      }
    },
  }));

  useEffect(() => {
    const body = bodyRef.current;
    if (!body) return;
    const update = () => {
      setWin({
        top: body.scrollTop,
        left: body.scrollLeft,
        width: body.clientWidth,
        height: body.clientHeight,
      });
    };
    update();
    const observer = new ResizeObserver(update);
    observer.observe(body);
    return () => observer.disconnect();
  }, [props.sheetId]);

  useLayoutEffect(() => {
    const node = editorNode.current;
    if (!node || !props.editing) return;
    if (document.activeElement !== node && node.value !== props.editing.draft) node.value = props.editing.draft;
  }, [props.editing]);

  useLayoutEffect(() => {
    if (!props.traceToken || props.reducedMotion) {
      setLines([]);
      return;
    }
    const shell = rootRef.current;
    if (!shell) return;
    const shellBox = shell.getBoundingClientRect();
    const target = shell.querySelector(`[data-addr="${props.active.r},${props.active.c}"]`);
    if (!target) return;
    const to = target.getBoundingClientRect();
    const next = props.trace.flatMap((cell) => {
      const from = shell.querySelector(`[data-addr="${cell.r},${cell.c}"]`)?.getBoundingClientRect();
      if (!from) return [];
      return [{
        x1: from.left + from.width / 2 - shellBox.left,
        y1: from.top + from.height / 2 - shellBox.top,
        x2: to.left + to.width / 2 - shellBox.left,
        y2: to.top + to.height / 2 - shellBox.top,
      }];
    });
    setLines(next);
  }, [props.traceToken, props.trace, props.active, props.reducedMotion, win]);

  function forwardWheel(event: ReactWheelEvent<HTMLDivElement>) {
    const body = bodyRef.current;
    if (!body || event.currentTarget === body) return;
    body.scrollTop += event.deltaY;
    body.scrollLeft += event.deltaX;
  }

  function syncScroll() {
    const body = bodyRef.current;
    if (!body) return;
    if (topRef.current) topRef.current.scrollLeft = body.scrollLeft;
    if (colRef.current) colRef.current.scrollLeft = body.scrollLeft;
    if (leftRef.current) leftRef.current.scrollTop = body.scrollTop;
    if (rowRef.current) rowRef.current.scrollTop = body.scrollTop;
    setWin((prev) => ({ ...prev, top: body.scrollTop, left: body.scrollLeft }));
  }

  function cellFromPointer(event: ReactPointerEvent, originR: number, originC: number, scroller: HTMLElement) {
    const rect = scroller.getBoundingClientRect();
    const x = event.clientX - rect.left + scroller.scrollLeft;
    const y = event.clientY - rect.top + scroller.scrollTop;
    return {
      r: rows.indexAt(rows.offset(originR) + Math.max(0, y)),
      c: cols.indexAt(cols.offset(originC) + Math.max(0, x)),
    };
  }

  function onPaneDown(event: ReactPointerEvent<HTMLDivElement>, originR: number, originC: number) {
    if (event.button !== 0) return;
    const hit = cellFromPointer(event, originR, originC, event.currentTarget);
    if ((event.target as HTMLElement).dataset.handle === 'fill') {
      drag.current = 'fill';
      event.currentTarget.setPointerCapture(event.pointerId);
      return;
    }
    if (event.pointerType !== 'touch') {
      drag.current = 'select';
      event.currentTarget.setPointerCapture(event.pointerId);
    }
    props.onSelect(hit.r, hit.c, event.shiftKey);
  }

  function onPaneMove(event: ReactPointerEvent<HTMLDivElement>, originR: number, originC: number) {
    if (!drag.current) return;
    const hit = cellFromPointer(event, originR, originC, event.currentTarget);
    if (drag.current === 'select') props.onSelect(hit.r, hit.c, true);
    if (drag.current === 'fill') {
      const area = normalizeRange(props.selection);
      const vertical = Math.abs(hit.r - area.r2) >= Math.abs(hit.c - area.c2);
      setFillTo(vertical
        ? { r1: Math.min(area.r1, hit.r), c1: area.c1, r2: Math.max(area.r2, hit.r), c2: area.c2 }
        : { r1: area.r1, c1: Math.min(area.c1, hit.c), r2: area.r2, c2: Math.max(area.c2, hit.c) });
    }
  }

  function onPaneUp() {
    if (drag.current === 'fill' && fillTo) props.onFill(fillTo);
    drag.current = null;
    setFillTo(null);
  }

  const bodyRows = visibleSpan(rows, freezeR, win.top, win.height);
  const bodyCols = visibleSpan(cols, freezeC, win.left, win.width);
  const selection = normalizeRange(fillTo ?? props.selection);

  function renderCells(r0: number, r1: number, c0: number, c1: number, originR: number, originC: number) {
    const nodes: ReactNode[] = [];
    for (let r = r0; r <= r1; r += 1) {
      if (rows.size(r) === 0) continue;
      for (let c = c0; c <= c1; c += 1) {
        if (cols.size(c) === 0) continue;
        const merge = findMerge(sheet, r, c);
        if (merge && (merge.r1 !== r || merge.c1 !== c)) continue;
        if (merge && (merge.r1 < originR || merge.c1 < originC)) continue;
        const r2 = merge?.r2 ?? r;
        const c2 = merge?.c2 ?? c;
        const record = sheet.cells[keyOf(r, c)];
        const paint = props.session.paint(props.sheetId, r, c, bounds);
        const selected = rangeContains(selection, r, c);
        const isActive = props.active.r === r && props.active.c === c;
        const traced = props.trace.some((cell) => cell.r === r && cell.c === c);
        const validation = props.session.validationAt(props.sheetId, r, c);
        const display = props.session.display(props.sheetId, r, c);
        const invalid = validation && display !== '' && !validation.values.includes(display);
        const style = cellStyle(record?.style, paint, isActive);
        nodes.push(
          <div
            key={keyOf(r, c)}
            role="gridcell"
            aria-rowindex={r + 1}
            aria-colindex={c + 1}
            aria-selected={selected}
            data-addr={`${r},${c}`}
            className={`cell${selected ? ' is-selected' : ''}${isActive ? ' is-active' : ''}${traced ? ' is-trace' : ''}${invalid ? ' is-invalid' : ''}${props.findHit?.r === r && props.findHit?.c === c ? ' is-find' : ''}`}
            style={{
              left: cols.offset(c) - cols.offset(originC),
              top: rows.offset(r) - rows.offset(originR),
              width: Math.max(cols.span(c, c2), cols.size(c)),
              height: Math.max(rows.span(r, r2), rows.size(r)),
              ...style,
            }}
            onDoubleClick={(event) => {
              event.stopPropagation();
              props.onEditStart(r, c);
            }}
          >
            {display}
            {sheet.filter && r === sheet.filter.headerRow && c >= sheet.filter.range.c1 && c <= sheet.filter.range.c2 ? (
              <button
                type="button"
                className="filter-mark"
                aria-label={`Filter ${colToA(c)}`}
                onPointerDown={(event) => event.stopPropagation()}
                onClick={(event) => props.onFilter(c, event.clientX, event.clientY)}
              />
            ) : null}
          </div>,
        );
      }
    }
    return nodes;
  }

  function overlay(originR: number, originC: number, showHandle: boolean) {
    const n = selection;
    if (n.r2 < originR || n.c2 < originC) return null;
    const r1 = Math.max(n.r1, originR);
    const c1 = Math.max(n.c1, originC);
    const style: CSSProperties = {
      left: cols.offset(c1) - cols.offset(originC),
      top: rows.offset(r1) - rows.offset(originR),
      width: cols.offset(n.c2 + 1) - cols.offset(c1),
      height: rows.offset(n.r2 + 1) - rows.offset(r1),
    };
    return (
      <div className={`selection${fillTo ? ' is-fill' : ''}`} style={style}>
        {showHandle ? <span className="fill-handle" data-handle="fill" /> : null}
      </div>
    );
  }

  function editor(originR: number, originC: number, endR: number, endC: number) {
    if (!props.editing) return null;
    const { r, c } = props.editing;
    if (r < originR || c < originC || r >= endR || c >= endC) return null;
    return (
      <textarea
        key={`${r},${c}`}
        className="cell-editor"
        aria-label="Cell editor"
        defaultValue={props.editing.draft}
        ref={(node) => {
          editorNode.current = node;
          if (!node || node.dataset.armed === '1') return;
          node.dataset.armed = '1';
          if (props.editing?.focus) {
            node.focus();
            const end = node.value.length;
            node.setSelectionRange(end, end);
          }
        }}
        style={{
          left: cols.offset(c) - cols.offset(originC),
          top: rows.offset(r) - rows.offset(originR),
          minWidth: cols.size(c),
          minHeight: rows.size(r),
        }}
        onChange={(event) => props.onEditDraft(event.target.value)}
        onKeyDown={(event) => {
          event.stopPropagation();
          if (event.key === 'Enter' && !event.altKey) {
            event.preventDefault();
            props.onCommit(event.shiftKey ? 'up' : 'down');
          } else if (event.key === 'Tab') {
            event.preventDefault();
            props.onCommit(event.shiftKey ? 'left' : 'right');
          } else if (event.key === 'Escape') {
            event.preventDefault();
            props.onCancel();
          }
        }}
        onBlur={(event) => {
          const next = event.relatedTarget as HTMLElement | null;
          if (next?.id === 'formula-input' || next?.classList.contains('name-box')) return;
          props.onCommit('stay');
        }}
      />
    );
  }

  function headers(axis: 'col' | 'row', start: number, end: number, origin: number) {
    const nodes: ReactNode[] = [];
    const index = axis === 'col' ? cols : rows;
    for (let i = start; i <= end; i += 1) {
      if (index.size(i) === 0) continue;
      const selected = axis === 'col'
        ? props.selection.c1 <= i && i <= props.selection.c2
        : props.selection.r1 <= i && i <= props.selection.r2;
      nodes.push(
        <div
          key={i}
          role={axis === 'col' ? 'columnheader' : 'rowheader'}
          className={`axis${selected ? ' is-in' : ''}`}
          style={axis === 'col'
            ? { left: index.offset(i) - index.offset(origin), width: index.size(i), height: COL_HEADER_HEIGHT }
            : { top: index.offset(i) - index.offset(origin), height: index.size(i), width: ROW_HEADER_WIDTH }}
        >
          {axis === 'col' ? colToA(i) : i + 1}
          <span
            className="resizer"
            onPointerDown={(event) => {
              event.stopPropagation();
              props.onGestureStart();
              const startPx = axis === 'col' ? event.clientX : event.clientY;
              const startSize = index.size(i);
              function move(next: PointerEvent) {
                const delta = (axis === 'col' ? next.clientX : next.clientY) - startPx;
                if (axis === 'col') props.onResizeCol(i, startSize + delta);
                else props.onResizeRow(i, startSize + delta);
              }
              function up() {
                props.onGestureEnd();
                window.removeEventListener('pointermove', move);
                window.removeEventListener('pointerup', up);
              }
              window.addEventListener('pointermove', move);
              window.addEventListener('pointerup', up);
            }}
          />
        </div>,
      );
    }
    return nodes;
  }

  const sweep = props.sweep;
  return (
    <div
      className="grid-shell"
      role="grid"
      aria-rowcount={ROW_COUNT}
      aria-colcount={COL_COUNT}
      aria-label={`${sheet.name}, ${ROW_COUNT.toLocaleString('en-US')} rows by ${COL_COUNT} columns`}
      tabIndex={0}
      ref={rootRef}
      style={{
        gridTemplateColumns: `${ROW_HEADER_WIDTH}px ${frozenW}px minmax(0,1fr)`,
        gridTemplateRows: `${COL_HEADER_HEIGHT}px ${frozenH}px minmax(0,1fr)`,
      }}
      onKeyDown={props.onKeyDown}
      onCopy={props.onCopy}
      onCut={props.onCut}
      onPaste={props.onPaste}
      onContextMenu={(event) => {
        const body = bodyRef.current;
        if (!body) return;
        event.preventDefault();
        const hit = cellFromPointer(event as unknown as ReactPointerEvent, freezeR, freezeC, body);
        props.onContext(hit.r, hit.c, event.clientX, event.clientY);
      }}
    >
      <div className="pane corner">
        <button type="button" className="select-all" aria-label="Select the used range" onClick={() => props.onSelect(0, 0, false)} />
      </div>
      <div className="pane head-x">
        <div className="pane-inner" style={{ width: frozenW, height: COL_HEADER_HEIGHT }}>
          {headers('col', 0, Math.max(0, freezeC - 1), 0)}
        </div>
      </div>
      <div className="pane head-x scroll-x" ref={colRef}>
        <div className="pane-inner" style={{ width: cols.total() - frozenW, height: COL_HEADER_HEIGHT }}>
          {headers('col', bodyCols.start, bodyCols.end, freezeC)}
        </div>
      </div>
      <div className="pane head-y">
        <div className="pane-inner" style={{ width: ROW_HEADER_WIDTH, height: frozenH }}>
          {headers('row', 0, Math.max(0, freezeR - 1), 0)}
        </div>
      </div>
      <div
        className="pane cells span-ok"
        onWheel={forwardWheel}
        onPointerDown={(event) => onPaneDown(event, 0, 0)}
        onPointerMove={(event) => onPaneMove(event, 0, 0)}
        onPointerUp={onPaneUp}
      >
        <div className="pane-inner" style={{ width: frozenW, height: frozenH }}>
          {renderCells(0, Math.max(0, freezeR - 1), 0, Math.max(0, freezeC - 1), 0, 0)}
          {overlay(0, 0, selection.r2 < freezeR && selection.c2 < freezeC)}
          {editor(0, 0, freezeR, freezeC)}
        </div>
      </div>
      <div
        className="pane cells scroll-x"
        ref={topRef}
        onWheel={forwardWheel}
        onPointerDown={(event) => onPaneDown(event, 0, freezeC)}
        onPointerMove={(event) => onPaneMove(event, 0, freezeC)}
        onPointerUp={onPaneUp}
      >
        <div className="pane-inner" style={{ width: cols.total() - frozenW, height: frozenH }}>
          {renderCells(0, Math.max(0, freezeR - 1), bodyCols.start, bodyCols.end, 0, freezeC)}
          {overlay(0, freezeC, selection.r2 < freezeR && selection.c2 >= freezeC)}
          {editor(0, freezeC, freezeR, COL_COUNT)}
        </div>
      </div>
      <div className="pane head-y scroll-y" ref={rowRef}>
        <div className="pane-inner" style={{ width: ROW_HEADER_WIDTH, height: rows.total() - frozenH }}>
          {headers('row', bodyRows.start, bodyRows.end, freezeR)}
        </div>
      </div>
      <div
        className="pane cells scroll-y"
        ref={leftRef}
        onWheel={forwardWheel}
        onPointerDown={(event) => onPaneDown(event, freezeR, 0)}
        onPointerMove={(event) => onPaneMove(event, freezeR, 0)}
        onPointerUp={onPaneUp}
      >
        <div className="pane-inner" style={{ width: frozenW, height: rows.total() - frozenH }}>
          {renderCells(bodyRows.start, bodyRows.end, 0, Math.max(0, freezeC - 1), freezeR, 0)}
          {overlay(freezeR, 0, selection.r2 >= freezeR && selection.c2 < freezeC)}
          {editor(freezeR, 0, ROW_COUNT, freezeC)}
        </div>
      </div>
      <div
        className="pane cells body"
        ref={bodyRef}
        onScroll={syncScroll}
        onPointerDown={(event) => onPaneDown(event, freezeR, freezeC)}
        onPointerMove={(event) => onPaneMove(event, freezeR, freezeC)}
        onPointerUp={onPaneUp}
      >
        <div className="pane-inner" style={{ width: Math.max(1, cols.total() - frozenW), height: Math.max(1, rows.total() - frozenH) }}>
          {renderCells(bodyRows.start, bodyRows.end, bodyCols.start, bodyCols.end, freezeR, freezeC)}
          {overlay(freezeR, freezeC, selection.r2 >= freezeR && selection.c2 >= freezeC)}
          {editor(freezeR, freezeC, ROW_COUNT, COL_COUNT)}
          {sweep && !props.reducedMotion ? (
            <div
              key={props.sweepToken}
              className="sweep"
              style={{
                left: cols.offset(sweep.c1) - cols.offset(freezeC),
                top: rows.offset(sweep.r1) - rows.offset(freezeR),
                width: cols.span(sweep.c1, sweep.c2),
                height: rows.span(sweep.r1, sweep.r2),
              }}
            />
          ) : null}
          {sheet.charts.map((chart) => (
            <ChartCard
              key={chart.id}
              chart={chart}
              session={props.session}
              sheetId={props.sheetId}
              originX={frozenW}
              originY={frozenH}
              onMove={props.onChartMove}
              onRemove={props.onChartRemove}
            />
          ))}
        </div>
      </div>
      {lines.length && !props.reducedMotion ? (
        <svg className="trace" aria-hidden="true">
          {lines.map((line, index) => (
            <line key={`${props.traceToken}-${index}`} x1={line.x1} y1={line.y1} x2={line.x2} y2={line.y2} />
          ))}
        </svg>
      ) : null}
    </div>
  );
});

function cellStyle(
  style: CellStyle | undefined,
  paint: { fill?: string; color?: string },
  active: boolean,
): CSSProperties {
  const border = style?.border;
  return {
    fontWeight: style?.bold ? 650 : 450,
    fontStyle: style?.italic ? 'italic' : 'normal',
    textDecoration: [style?.underline ? 'underline' : '', style?.strike ? 'line-through' : ''].filter(Boolean).join(' ') || undefined,
    fontFamily: style?.fontFamily === 'Syne' ? 'var(--display)' : style?.fontFamily === 'mono' ? 'var(--mono)' : undefined,
    fontSize: style?.fontSize ? `${style.fontSize}px` : undefined,
    color: paint.color || adaptiveColor(style?.color, paint.fill || style?.fill),
    background: active ? 'var(--paper-strong)' : paint.fill || style?.fill || undefined,
    textAlign: style?.align,
    alignItems: style?.valign === 'top' ? 'flex-start' : style?.valign === 'bottom' ? 'flex-end' : 'center',
    whiteSpace: style?.wrap ? 'normal' : 'nowrap',
    borderTop: border?.top ? `${weight(border.top.style)} solid ${border.top.color}` : undefined,
    borderRight: border?.right ? `${weight(border.right.style)} solid ${border.right.color}` : undefined,
    borderBottom: border?.bottom ? `${weight(border.bottom.style)} solid ${border.bottom.color}` : undefined,
    borderLeft: border?.left ? `${weight(border.left.style)} solid ${border.left.color}` : undefined,
  };
}

function adaptiveColor(color: string | undefined, fill: string | undefined): string | undefined {
  if (!color) return undefined;
  const value = color.toLowerCase();
  const ink = value === '#1b3148' || value === '#1c2430' || value === '#000000';
  const muted = value === '#3d4a5c' || value === '#5e6b7c';
  if ((ink || muted) && !isLight(fill)) return muted ? 'var(--muted)' : undefined;
  return color;
}

function isLight(color: string | undefined): boolean {
  if (!color || !color.startsWith('#')) return false;
  const hex = color.slice(1);
  const full = hex.length === 3 ? hex.split('').map((ch) => ch + ch).join('') : hex;
  const r = Number.parseInt(full.slice(0, 2), 16);
  const g = Number.parseInt(full.slice(2, 4), 16);
  const b = Number.parseInt(full.slice(4, 6), 16);
  if ([r, g, b].some((part) => Number.isNaN(part))) return false;
  return (r * 299 + g * 587 + b * 114) / 1000 > 160;
}

function weight(style: 'thin' | 'medium' | 'thick'): string {
  if (style === 'medium') return '2px';
  if (style === 'thick') return '3px';
  return '1px';
}
