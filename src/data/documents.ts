import { cloneData } from '../domain/types';
import { getMeta, getWorkbook, setMeta } from './workbooks';
import { saveLinkedRange } from './linkedRanges';
import { officeDb } from './db';
import { WorkbookSession } from '../engine/session';
import { BRIEF_ID, PIPELINE_LINK_ID, createBriefing } from '../rooms/docs/sample';
import type { DocModel } from '../rooms/docs/model';

export async function listDocuments(): Promise<DocModel[]> {
  const db = await officeDb();
  const all = await db.getAll('documents');
  return all.sort((a, b) => b.updatedAt - a.updatedAt);
}

export async function saveDocument(model: DocModel): Promise<void> {
  const db = await officeDb();
  await db.put('documents', cloneData(model));
}

export async function deleteDocument(id: string): Promise<void> {
  const db = await officeDb();
  await db.delete('documents', id);
}

let docBoot: Promise<DocModel[]> | null = null;

export function loadDocuments(): Promise<DocModel[]> {
  if (!docBoot) docBoot = bootDocuments();
  return docBoot.then((docs) => docs.map((doc) => cloneData(doc)));
}

export function resetDocumentCache(): void {
  docBoot = null;
}

async function bootDocuments(): Promise<DocModel[]> {
  const existing = await listDocuments();
  const seeded = await getMeta('docSeeded');
  if (!seeded && !existing.some((doc) => doc.id === BRIEF_ID)) {
    await ensurePipelineLink();
    const brief = createBriefing();
    await saveDocument(brief);
    await setMeta('docSeeded', '1');
    return [brief, ...existing];
  }
  if (!seeded) await setMeta('docSeeded', '1');
  return existing;
}

async function ensurePipelineLink(): Promise<void> {
  const book = await getWorkbook('book_desk_plan');
  if (!book) return;
  const session = new WorkbookSession(book);
  const sheet = session.model.sheets.find((item) => item.id === 'sheet_pipeline') ?? session.activeSheet();
  const cells = session.snapshotRange(sheet.id, { r1: 3, c1: 0, r2: 10, c2: 7 });
  session.destroy();
  const now = Date.now();
  await saveLinkedRange({
    id: PIPELINE_LINK_ID,
    workbookId: book.id,
    sheetId: sheet.id,
    sheetName: sheet.name,
    a1: 'A4:H11',
    r1: 3,
    c1: 0,
    r2: 10,
    c2: 7,
    cells,
    createdAt: now,
    updatedAt: now,
  });
}
