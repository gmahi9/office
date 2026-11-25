import type { PlainValue } from '../domain/types';
import { officeDb } from './db';

/**
 * A linked range is a sheet selection published for another room.
 * Sheets writes the snapshot. Docs reads it. The values are the last save,
 * so a reader does not need the formula engine.
 */
export interface LinkedRangeCell {
  input: string | null;
  value: PlainValue;
  display: string;
}

export interface LinkedRange {
  id: string;
  workbookId: string;
  sheetId: string;
  sheetName: string;
  a1: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  cells: LinkedRangeCell[][];
  createdAt: number;
  updatedAt: number;
}

type LinkListener = (range: LinkedRange) => void;
const linkListeners = new Set<LinkListener>();

/** Same-tab notice so an open document can refresh a linked table. */
export function subscribeLinkedRanges(listener: LinkListener): () => void {
  linkListeners.add(listener);
  return () => linkListeners.delete(listener);
}

export async function saveLinkedRange(range: LinkedRange): Promise<void> {
  const db = await officeDb();
  await db.put('links', range);
  for (const listener of linkListeners) listener(range);
}

export async function getLinkedRange(id: string): Promise<LinkedRange | undefined> {
  const db = await officeDb();
  return db.get('links', id);
}

export async function listLinkedRanges(workbookId?: string): Promise<LinkedRange[]> {
  const db = await officeDb();
  if (workbookId) return db.getAllFromIndex('links', 'byWorkbook', workbookId);
  return db.getAll('links');
}

export async function deleteLinkedRange(id: string): Promise<void> {
  const db = await officeDb();
  await db.delete('links', id);
}

export async function deleteLinkedRangesForWorkbook(workbookId: string): Promise<void> {
  const ranges = await listLinkedRanges(workbookId);
  const db = await officeDb();
  const tx = db.transaction('links', 'readwrite');
  await Promise.all(ranges.map((range) => tx.store.delete(range.id)));
  await tx.done;
}
