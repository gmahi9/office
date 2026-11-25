import { openDB, type DBSchema, type IDBPDatabase } from 'idb';
import type { WorkbookModel } from '../domain/types';
import type { DocModel } from '../rooms/docs/model';
import type { LinkedRange } from './linkedRanges';
import type { UsageStat } from './usage';

interface OfficeSchema extends DBSchema {
  workbooks: { key: string; value: WorkbookModel };
  documents: { key: string; value: DocModel };
  links: { key: string; value: LinkedRange; indexes: { byWorkbook: string } };
  usage: { key: string; value: UsageStat };
  meta: { key: string; value: string };
}

const DB_NAME = 'aitek-office';

let dbPromise: Promise<IDBPDatabase<OfficeSchema>> | null = null;

export function officeDb(): Promise<IDBPDatabase<OfficeSchema>> {
  if (!dbPromise) {
    dbPromise = openDB<OfficeSchema>(DB_NAME, 2, {
      upgrade(db) {
        if (!db.objectStoreNames.contains('workbooks')) db.createObjectStore('workbooks', { keyPath: 'id' });
        if (!db.objectStoreNames.contains('links')) {
          const links = db.createObjectStore('links', { keyPath: 'id' });
          links.createIndex('byWorkbook', 'workbookId');
        }
        if (!db.objectStoreNames.contains('usage')) db.createObjectStore('usage', { keyPath: 'commandId' });
        if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta');
        if (!db.objectStoreNames.contains('documents')) db.createObjectStore('documents', { keyPath: 'id' });
      },
    });
  }
  return dbPromise;
}

export async function resetOfficeDb(): Promise<void> {
  if (dbPromise) {
    const db = await dbPromise;
    db.close();
    dbPromise = null;
  }
  await new Promise<void>((resolve, reject) => {
    const request = indexedDB.deleteDatabase(DB_NAME);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
    request.onblocked = () => resolve();
  });
}
