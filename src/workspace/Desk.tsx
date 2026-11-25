import { useRef, type PointerEvent as ReactPointerEvent, type ReactNode, type WheelEvent } from 'react';
import { SOURCE_URL } from '../legal/About';

export interface DeskCard {
  id: string;
  kind: 'sheet' | 'doc';
  title: string;
  x: number;
  y: number;
  preview: ReactNode;
}

export function Desk({
  cards,
  selected,
  shredId,
  onSelect,
  onOpen,
  onMove,
  onDelete,
  onNewSheet,
  onNewDoc,
  onAbout,
}: {
  cards: DeskCard[];
  selected: string | null;
  shredId: string | null;
  onSelect: (id: string) => void;
  onOpen: (id: string, rect: DOMRect) => void;
  onMove: (id: string, x: number, y: number) => void;
  onDelete: (id: string) => void;
  onNewSheet: () => void;
  onNewDoc: () => void;
  onAbout: () => void;
}) {
  const world = useRef<HTMLDivElement>(null);
  const cam = useRef({ x: 16, y: 20, scale: 1 });
  const frame = useRef<HTMLDivElement>(null);

  function apply() {
    const node = world.current;
    if (!node) return;
    node.style.transform = `translate(${cam.current.x}px, ${cam.current.y}px) scale(${cam.current.scale})`;
  }

  function onWheel(event: WheelEvent<HTMLDivElement>) {
    event.preventDefault();
    const rect = event.currentTarget.getBoundingClientRect();
    const mx = event.clientX - rect.left;
    const my = event.clientY - rect.top;
    const next = clamp(cam.current.scale * (event.deltaY < 0 ? 1.08 : 0.92), 0.42, 1.35);
    const wx = (mx - cam.current.x) / cam.current.scale;
    const wy = (my - cam.current.y) / cam.current.scale;
    cam.current = { scale: next, x: mx - wx * next, y: my - wy * next };
    apply();
  }

  function onBackgroundDown(event: ReactPointerEvent<HTMLDivElement>) {
    if ((event.target as HTMLElement).closest('.room-card')) return;
    const startX = event.clientX;
    const startY = event.clientY;
    const origin = { ...cam.current };
    function move(next: PointerEvent) {
      cam.current = { ...cam.current, x: origin.x + next.clientX - startX, y: origin.y + next.clientY - startY };
      apply();
    }
    function up() {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    }
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }

  return (
    <div className="desk" ref={frame} onWheel={onWheel} onPointerDown={onBackgroundDown}>
      <div className="desk-tools">
        <button type="button" className="tool" onClick={onNewSheet}>New workbook</button>
        <button type="button" className="tool" onClick={onNewDoc}>New document</button>
        <button
          type="button"
          className="tool"
          onClick={() => {
            cam.current = { x: 16, y: 20, scale: cards.length > 2 ? 0.7 : 1 };
            apply();
          }}
        >
          See the desk
        </button>
      </div>
      {cards.length === 0 ? (
        <div className="desk-empty">
          <h2>The desk is clear</h2>
          <p>Files stay in this browser. Start a workbook or a document.</p>
          <button type="button" className="progress-btn idle-btn" onClick={onNewSheet}>New workbook</button>
        </div>
      ) : null}
      <div className="desk-world" ref={world} style={{ transform: 'translate(16px, 20px) scale(1)' }}>
        {cards.map((card) => {
          return (
            <article
              key={card.id}
              className={`room-card${selected === card.id ? ' is-selected' : ''}${shredId === card.id ? ' is-shred' : ''}`}
              style={{ left: card.x, top: card.y }}
              data-room={card.id}
              onPointerDown={(event) => {
                if ((event.target as HTMLElement).closest('button')) return;
                event.stopPropagation();
                onSelect(card.id);
                const startX = event.clientX;
                const startY = event.clientY;
                const originX = card.x;
                const originY = card.y;
                let moved = false;
                function move(next: PointerEvent) {
                  const dx = next.clientX - startX;
                  const dy = next.clientY - startY;
                  if (Math.hypot(dx, dy) > 4) moved = true;
                  onMove(card.id, originX + dx / cam.current.scale, originY + dy / cam.current.scale);
                }
                function up() {
                  window.removeEventListener('pointermove', move);
                  window.removeEventListener('pointerup', up);
                  if (!moved) onSelect(card.id);
                }
                window.addEventListener('pointermove', move);
                window.addEventListener('pointerup', up);
              }}
              onDoubleClick={(event) => {
                const rect = (event.currentTarget as HTMLElement).getBoundingClientRect();
                onOpen(card.id, rect);
              }}
            >
              <header>
                <span className="kind">{card.kind === 'doc' ? 'Doc' : 'Sheet'}</span>
                <h3>{card.title}</h3>
              </header>
              {card.preview}
              <footer>
                <button type="button" onClick={() => {
                  const node = document.querySelector(`[data-room="${card.id}"]`);
                  onOpen(card.id, node?.getBoundingClientRect() ?? new DOMRect(80, 80, 280, 180));
                }}>Open</button>
                <button type="button" onClick={() => onDelete(card.id)}>Delete</button>
              </footer>
            </article>
          );
        })}
      </div>
      <p className="legal-bar">
        <a href={SOURCE_URL}>Open source</a>
        <button type="button" className="text-btn" onClick={onAbout}>Licences</button>
      </p>
    </div>
  );
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}
