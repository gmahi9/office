import { COL_COUNT, ROW_COUNT } from './types';

export function jumpCell(
  isEmpty: (r: number, c: number) => boolean,
  isHidden: (r: number, c: number) => boolean,
  r: number,
  c: number,
  dr: number,
  dc: number,
): { r: number; c: number } {
  const step = (row: number, col: number) => ({ r: row + dr, c: col + dc });
  const inside = (row: number, col: number) => row >= 0 && col >= 0 && row < ROW_COUNT && col < COL_COUNT;
  let nr = r;
  let nc = c;
  const first = step(nr, nc);
  if (!inside(first.r, first.c)) return { r, c };
  nr = first.r;
  nc = first.c;
  while (inside(nr, nc) && isHidden(nr, nc)) {
    const nxt = step(nr, nc);
    if (!inside(nxt.r, nxt.c)) return { r: nr, c: nc };
    nr = nxt.r;
    nc = nxt.c;
  }
  const startEmpty = isEmpty(r, c);
  const landedEmpty = isEmpty(nr, nc);
  if (startEmpty || landedEmpty) {
    while (inside(nr, nc) && (isEmpty(nr, nc) || isHidden(nr, nc))) {
      const nxt = step(nr, nc);
      if (!inside(nxt.r, nxt.c)) return { r: nr, c: nc };
      if (!isEmpty(nxt.r, nxt.c) && !isHidden(nxt.r, nxt.c)) return nxt;
      nr = nxt.r;
      nc = nxt.c;
    }
    return { r: nr, c: nc };
  }
  while (inside(nr, nc)) {
    const nxt = step(nr, nc);
    if (!inside(nxt.r, nxt.c) || isEmpty(nxt.r, nxt.c)) return { r: nr, c: nc };
    if (isHidden(nxt.r, nxt.c)) {
      nr = nxt.r;
      nc = nxt.c;
      continue;
    }
    nr = nxt.r;
    nc = nxt.c;
  }
  return { r: nr, c: nc };
}
