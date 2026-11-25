import { cloneData, type WorkbookModel } from '../domain/types';
import { createSampleWorkbook } from '../domain/sample';
import { officeDb } from './db';

export async function listWorkbooks(): Promise<WorkbookModel[]> {
  const db = await officeDb();
  const all = await db.getAll('workbooks');
  return all.sort((a, b) => b.updatedAt - a.updatedAt);
}

export async function getWorkbook(id: string): Promise<WorkbookModel | undefined> {
  const db = await officeDb();
  return db.get('workbooks', id);
}

export async function saveWorkbook(model: WorkbookModel): Promise<void> {
  const db = await officeDb();
  await db.put('workbooks', cloneData(model));
}

export async function deleteWorkbook(id: string): Promise<void> {
  const db = await officeDb();
  await db.delete('workbooks', id);
}

export async function getMeta(key: string): Promise<string | undefined> {
  const db = await officeDb();
  return db.get('meta', key);
}

export async function setMeta(key: string, value: string): Promise<void> {
  const db = await officeDb();
  await db.put('meta', value, key);
}

let bootPromise: Promise<WorkbookModel[]> | null = null;

/** Load the desk. The first visit on a browser plants one sample workbook. */
export function loadDesk(): Promise<WorkbookModel[]> {
  if (!bootPromise) bootPromise = bootDesk();
  return bootPromise.then((books) => books.map((book) => cloneData(book)));
}

const SAMPLE_VERSION = '2';

async function bootDesk(): Promise<WorkbookModel[]> {
  const existing = await listWorkbooks();
  const seeded = await getMeta('seeded');
  if (!seeded && existing.length === 0) {
    const sample = createSampleWorkbook();
    await saveWorkbook(sample);
    await setMeta('seeded', '1');
    await setMeta('sampleVersion', SAMPLE_VERSION);
    return [sample];
  }
  if (!seeded) await setMeta('seeded', '1');
  if ((await getMeta('sampleVersion')) !== SAMPLE_VERSION) {
    const current = existing.find((book) => book.id === 'book_desk_plan');
    if (current) {
      const fresh = createSampleWorkbook();
      fresh.createdAt = current.createdAt;
      await saveWorkbook(fresh);
      const index = existing.findIndex((book) => book.id === current.id);
      existing[index] = fresh;
    }
    await setMeta('sampleVersion', SAMPLE_VERSION);
  }
  return existing;
}

export function resetDeskCache(): void {
  bootPromise = null;
}
