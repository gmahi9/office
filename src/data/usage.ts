import { officeDb } from './db';

export interface UsageStat {
  commandId: string;
  count: number;
  lastUsed: number;
}

export async function readUsage(): Promise<UsageStat[]> {
  const db = await officeDb();
  return db.getAll('usage');
}

export async function recordUse(commandId: string, now = Date.now()): Promise<UsageStat[]> {
  const db = await officeDb();
  const prev = await db.get('usage', commandId);
  const next: UsageStat = {
    commandId,
    count: (prev?.count ?? 0) + 1,
    lastUsed: now,
  };
  await db.put('usage', next);
  return db.getAll('usage');
}

export async function resetUsage(): Promise<void> {
  const db = await officeDb();
  await db.clear('usage');
}
