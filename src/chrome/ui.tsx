import { useEffect, useId, useRef, useState, type CSSProperties, type ReactNode } from 'react';
import { ROOM_KINDS, type RoomKind } from '../rooms/types';

export function useReducedMotion(): boolean {
  const [reduced, setReduced] = useState(() => window.matchMedia('(prefers-reduced-motion: reduce)').matches);
  useEffect(() => {
    const query = window.matchMedia('(prefers-reduced-motion: reduce)');
    const onChange = () => setReduced(query.matches);
    query.addEventListener('change', onChange);
    return () => query.removeEventListener('change', onChange);
  }, []);
  return reduced;
}

export function Dialog({
  title,
  onClose,
  children,
}: {
  title: string;
  onClose: () => void;
  children: ReactNode;
}) {
  const titleId = useId();
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const previously = document.activeElement as HTMLElement | null;
    ref.current?.focus();
    const onKey = (event: KeyboardEvent) => {
      if (event.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => {
      window.removeEventListener('keydown', onKey);
      previously?.focus();
    };
  }, [onClose]);
  return (
    <div className="scrim" onMouseDown={onClose}>
      <div
        className="dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        tabIndex={-1}
        ref={ref}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="dialog-head">
          <h2 id={titleId}>{title}</h2>
          <button type="button" className="text-btn" onClick={onClose} aria-label="Close">
            Close
          </button>
        </header>
        {children}
      </div>
    </div>
  );
}

export function ProgressButton({
  label,
  doneLabel,
  onRun,
}: {
  label: string;
  doneLabel: string;
  onRun: (report: (value: number) => void) => Promise<void>;
}) {
  const [progress, setProgress] = useState(0);
  const [phase, setPhase] = useState<'idle' | 'run' | 'done'>('idle');
  async function click() {
    if (phase === 'run') return;
    setPhase('run');
    setProgress(0.08);
    try {
      await onRun((value) => setProgress(Math.max(0, Math.min(1, value))));
      setProgress(1);
      setPhase('done');
      window.setTimeout(() => {
        setPhase('idle');
        setProgress(0);
      }, 1100);
    } catch {
      setPhase('idle');
      setProgress(0);
    }
  }
  return (
    <button
      type="button"
      className={`progress-btn${phase === 'done' ? ' is-done' : ''}`}
      onClick={() => void click()}
      aria-busy={phase === 'run'}
      disabled={phase === 'run'}
    >
      <span className="progress-fill" style={{ width: `${progress * 100}%` }} />
      <span className="progress-label">
        {phase === 'done' ? (
          <span className="check" aria-hidden="true">
            <svg viewBox="0 0 16 16" width="14" height="14">
              <path d="M3 8.2 6.4 11.5 13 4.5" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
            </svg>
            {doneLabel}
          </span>
        ) : (
          label
        )}
      </span>
    </button>
  );
}

export interface ToastItem {
  id: number;
  message: string;
  undo?: () => void;
}

export function Toast({ item, onClose }: { item: ToastItem | null; onClose: () => void }) {
  if (!item) return null;
  return (
    <div className="toast" role="status">
      <span>{item.message}</span>
      {item.undo ? (
        <button
          type="button"
          onClick={() => {
            item.undo?.();
            onClose();
          }}
        >
          Undo
        </button>
      ) : null}
    </div>
  );
}

export function SuiteSwitch({
  active,
  onSoon,
  onChoose,
  spotlight,
}: {
  active: RoomKind;
  onSoon: (label: string) => void;
  onChoose?: (kind: RoomKind) => void;
  spotlight?: boolean;
}) {
  const index = Math.max(0, ROOM_KINDS.findIndex((room) => room.kind === active));
  return (
    <div className={spotlight ? 'suite spotlight-bar' : 'suite'} role="tablist" aria-label="Office">
      {spotlight ? <span className="spotlight" style={{ '--i': String(index) } as CSSProperties} /> : null}
      {ROOM_KINDS.map((room) => (
        <button
          key={room.kind}
          type="button"
          role="tab"
          aria-selected={room.kind === active}
          aria-disabled={!room.available}
          className={room.kind === active ? 'is-active' : ''}
          onClick={() => {
            if (!room.available) onSoon(room.label);
            else onChoose?.(room.kind);
          }}
        >
          <span>{room.label}</span>
          {!room.available ? <small>Soon</small> : null}
        </button>
      ))}
    </div>
  );
}
