import { useEffect, useRef, useState } from 'react';
import { SuiteSwitch, Toast, useReducedMotion, type ToastItem } from './chrome/ui';
import { cloneData, createWorkbook, type WorkbookModel } from './domain/types';
import { deleteLinkedRangesForWorkbook, listLinkedRanges, saveLinkedRange } from './data/linkedRanges';
import { deleteDocument, loadDocuments, saveDocument } from './data/documents';
import { readUsage, recordUse, resetUsage, type UsageStat } from './data/usage';
import { deleteWorkbook, loadDesk, saveWorkbook } from './data/workbooks';
import { WorkbookSession } from './engine/session';
import { About } from './legal/About';
import { DocRoom } from './rooms/docs/DocRoom';
import { emptyDocument, previewText, type DocModel } from './rooms/docs/model';
import { SheetRoom } from './rooms/sheets/SheetRoom';
import type { RoomKind } from './rooms/types';
import { Desk, type DeskCard } from './workspace/Desk';

const THEME_KEY = 'aitek-office-theme';

export default function App() {
  const reduced = useReducedMotion();
  const sessions = useRef(new Map<string, WorkbookSession>());
  const dirtyDocs = useRef(new Set<string>());
  const [books, setBooks] = useState<WorkbookModel[] | null>(null);
  const [docs, setDocs] = useState<DocModel[]>([]);
  const [storageNote, setStorageNote] = useState('');
  const [focus, setFocus] = useState<string | null>(null);
  const [selected, setSelected] = useState<string | null>(null);
  const [shredId, setShredId] = useState<string | null>(null);
  const [toast, setToast] = useState<ToastItem | null>(null);
  const [about, setAbout] = useState(false);
  const [usage, setUsage] = useState<UsageStat[]>([]);
  const [rev, setRev] = useState(0);
  const [theme, setTheme] = useState<'light' | 'dark'>(() => {
    const stored = localStorage.getItem(THEME_KEY);
    if (stored === 'dark' || stored === 'light') return stored;
    return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  });

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem(THEME_KEY, theme);
    const color = theme === 'dark' ? '#101820' : '#eef5fb';
    document.querySelector('meta[name="theme-color"]')?.setAttribute('content', color);
  }, [theme]);

  useEffect(() => {
    let gone = false;
    loadDesk()
      .then(async (loaded) => {
        const documents = await loadDocuments();
        if (gone) return;
        for (const book of loaded) sessions.current.set(book.id, new WorkbookSession(book));
        setBooks(loaded);
        setDocs(documents);
        setSelected(loaded[0]?.id ?? documents[0]?.id ?? null);
      })
      .catch(() => {
        if (gone) return;
        setStorageNote('This browser is not storing the desk. Files last until you close the tab.');
        const memory = createWorkbook('Desk plan');
        sessions.current.set(memory.id, new WorkbookSession(memory));
        setBooks([memory]);
      });
    void readUsage().then(setUsage).catch(() => undefined);
    return () => {
      gone = true;
    };
  }, []);

  useEffect(() => {
    if (!focus) return;
    const session = sessions.current.get(focus);
    const doc = docs.find((item) => item.id === focus);
    if (!session?.dirty && !(doc && dirtyDocs.current.has(doc.id))) return;
    const timer = window.setTimeout(() => {
      if (session?.dirty) {
        session.dirty = false;
        void saveWorkbook(session.model).catch(() => undefined);
        void refreshLinks(session);
        setBooks((current) => current?.map((book) => book.id === session.model.id
          ? { ...book, title: session.model.title, updatedAt: session.model.updatedAt, x: session.model.x, y: session.model.y }
          : book) ?? current);
      }
      if (doc && dirtyDocs.current.has(doc.id)) {
        dirtyDocs.current.delete(doc.id);
        void saveDocument(doc).catch(() => undefined);
      }
    }, 450);
    return () => window.clearTimeout(timer);
  }, [rev, focus, docs]);

  function ensure(model: WorkbookModel): WorkbookSession {
    let session = sessions.current.get(model.id);
    if (!session) {
      session = new WorkbookSession(model);
      sessions.current.set(model.id, session);
    }
    return session;
  }

  function openRoom(id: string, _rect: DOMRect) {
    setFocus(id);
    setSelected(id);
    setAbout(false);
  }

  function closeRoom() {
    setFocus(null);
  }

  function notify(message: string, undo?: () => void) {
    const item = { id: Date.now(), message, undo };
    setToast(item);
    window.setTimeout(() => setToast((current) => (current?.id === item.id ? null : current)), 6000);
  }

  function onUse(id: string) {
    if (id === 'reset-usage') {
      void resetUsage().then(() => setUsage([]));
      notify('Toolbar order reset.');
      return;
    }
    void recordUse(id).then(setUsage).catch(() => undefined);
  }

  function chooseKind(kind: RoomKind) {
    const pool = kind === 'doc'
      ? docs.map((doc) => ({ id: doc.id, updatedAt: doc.updatedAt }))
      : (books ?? []).map((book) => ({ id: book.id, updatedAt: book.updatedAt }));
    const latest = [...pool].sort((a, b) => b.updatedAt - a.updatedAt)[0];
    if (!latest) {
      notify(kind === 'doc' ? 'No document on the desk yet.' : 'No workbook on the desk yet.');
      return;
    }
    if (focus === latest.id) closeRoom();
    else openRoom(latest.id, new DOMRect());
  }

  async function addWorkbook() {
    const index = (books?.length ?? 0) + docs.length;
    const model = createWorkbook('Untitled', 28 + (index % 3) * 340, 64 + Math.floor(index / 3) * 250);
    ensure(model);
    await saveWorkbook(model).catch(() => undefined);
    setBooks((current) => [...(current ?? []), model]);
    openRoom(model.id, new DOMRect(120, 120, 280, 180));
  }

  async function addDocument() {
    const index = (books?.length ?? 0) + docs.length;
    const model = emptyDocument('Untitled', 28 + (index % 3) * 340, 64 + Math.floor(index / 3) * 250);
    await saveDocument(model).catch(() => undefined);
    setDocs((current) => [model, ...current]);
    openRoom(model.id, new DOMRect(160, 120, 280, 180));
  }

  function removeFile(id: string) {
    const session = sessions.current.get(id);
    const doc = docs.find((item) => item.id === id);
    if (!session && !doc) return;
    const title = session?.model.title ?? doc?.title ?? 'File';
    const snapshotBook = session ? cloneData(session.model) : null;
    const snapshotDoc = doc ? cloneData(doc) : null;
    setShredId(id);
    window.setTimeout(() => {
      if (session && snapshotBook) {
        session.destroy();
        sessions.current.delete(id);
        setBooks((current) => current?.filter((book) => book.id !== id) ?? current);
        void deleteWorkbook(id);
        void deleteLinkedRangesForWorkbook(id);
      }
      if (snapshotDoc) {
        setDocs((current) => current.filter((item) => item.id !== id));
        void deleteDocument(id);
      }
      setShredId(null);
      if (focus === id) setFocus(null);
      notify(`${title} left the desk.`, () => {
        if (snapshotBook) {
          sessions.current.set(snapshotBook.id, new WorkbookSession(snapshotBook));
          void saveWorkbook(snapshotBook);
          setBooks((current) => [...(current ?? []), snapshotBook]);
        }
        if (snapshotDoc) {
          void saveDocument(snapshotDoc);
          setDocs((current) => [snapshotDoc, ...current]);
        }
      });
    }, reduced ? 0 : 360);
  }

  const focusedBook = books?.find((book) => book.id === focus);
  const focusedSession = focusedBook ? sessions.current.get(focusedBook.id) : undefined;
  const focusedDoc = docs.find((doc) => doc.id === focus);
  const activeKind: RoomKind = focusedDoc ? 'doc' : 'sheet';
  const cards: DeskCard[] = [
    ...(books ?? []).map((book) => ({
      id: book.id,
      kind: 'sheet' as const,
      title: book.title,
      x: book.x,
      y: book.y,
      preview: <SheetPreview session={sessions.current.get(book.id)} />,
    })),
    ...docs.map((doc) => ({
      id: doc.id,
      kind: 'doc' as const,
      title: doc.title,
      x: doc.x,
      y: doc.y,
      preview: <p className="doc-preview">{previewText(doc) || 'Empty page'}</p>,
    })),
  ];

  return (
    <div className="app">
      <a className="skip" href="#stage">Skip to the desk</a>
      <header className="topbar">
        <div className="brand">
          <img src="./favicon.svg" alt="" width="28" height="28" />
          <strong>Office</strong>
          <span>{focusedDoc ? 'Docs' : focusedSession ? 'Sheets' : 'Desk'}</span>
        </div>
        <SuiteSwitch
          active={activeKind}
          onSoon={(label) => notify(`${label} is not in this build yet.`)}
          onChoose={chooseKind}
        />
        {focusedSession ? (
          <input
            className="title-input"
            aria-label="Workbook title"
            value={focusedSession.model.title}
            onChange={(event) => {
              focusedSession.model.title = event.target.value;
              focusedSession.dirty = true;
              setRev((value) => value + 1);
            }}
          />
        ) : focusedDoc ? (
          <input
            className="title-input"
            aria-label="Document title"
            value={focusedDoc.title}
            onChange={(event) => {
              const title = event.target.value;
              setDocs((current) => current.map((item) => item.id === focusedDoc.id ? { ...item, title } : item));
              dirtyDocs.current.add(focusedDoc.id);
              setRev((value) => value + 1);
            }}
          />
        ) : <p className="title-input quiet">Your desk</p>}
        <button type="button" className="tool" onClick={() => setTheme((value) => value === 'light' ? 'dark' : 'light')}>
          {theme === 'light' ? 'Dark' : 'Light'}
        </button>
        {focus ? <button type="button" className="tool" onClick={closeRoom}>Desk</button> : null}
      </header>
      {storageNote ? <p className="banner" role="status">{storageNote}</p> : null}
      <main className="stage" id="stage">
        {books === null ? <p className="loading">Opening the desk…</p> : (
          <Desk
            cards={cards}
            selected={selected}
            shredId={shredId}
            onSelect={setSelected}
            onOpen={openRoom}
            onMove={(id, x, y) => {
              const session = sessions.current.get(id);
              if (session) {
                session.model.x = x;
                session.model.y = y;
                session.dirty = true;
                setBooks((current) => current?.map((book) => book.id === id ? { ...book, x, y } : book) ?? current);
              } else {
                setDocs((current) => current.map((doc) => doc.id === id ? { ...doc, x, y } : doc));
                dirtyDocs.current.add(id);
              }
              setRev((value) => value + 1);
            }}
            onDelete={removeFile}
            onNewSheet={() => void addWorkbook()}
            onNewDoc={() => void addDocument()}
            onAbout={() => setAbout(true)}
          />
        )}
        {focusedSession ? (
          <div className={`room-layer${reduced ? '' : ' is-zooming'}`}>
            <SheetRoom
              key={focusedSession.model.id}
              session={focusedSession}
              usage={usage}
              reducedMotion={reduced}
              onUse={onUse}
              onDirty={() => setRev((value) => value + 1)}
              onToast={notify}
              onImported={(models) => {
                models.forEach((model, index) => {
                  model.x = 140 + index * 36;
                  model.y = 120 + index * 28;
                  ensure(model);
                  void saveWorkbook(model);
                });
                setBooks((current) => [...models, ...(current ?? [])]);
                const first = models[0];
                if (first) openRoom(first.id, new DOMRect(140, 120, 300, 200));
              }}
            />
          </div>
        ) : null}
        {focusedDoc ? (
          <div className={`room-layer${reduced ? '' : ' is-zooming'}`}>
            <DocRoom
              key={focusedDoc.id}
              doc={focusedDoc}
              usage={usage}
              reducedMotion={reduced}
              onUse={onUse}
              onToast={notify}
              onChange={(patch) => {
                setDocs((current) => current.map((item) => item.id === focusedDoc.id
                  ? { ...item, ...patch, updatedAt: Date.now() }
                  : item));
                dirtyDocs.current.add(focusedDoc.id);
                setRev((value) => value + 1);
              }}
            />
          </div>
        ) : null}
      </main>
      <SuiteSwitch
        active={activeKind}
        spotlight
        onSoon={(label) => notify(`${label} is not in this build yet.`)}
        onChoose={chooseKind}
      />
      {about ? <About onClose={() => setAbout(false)} /> : null}
      <Toast item={toast} onClose={() => setToast(null)} />
    </div>
  );
}

function SheetPreview({ session }: { session: WorkbookSession | undefined }) {
  if (!session) return <p className="preview-wait">Opening…</p>;
  const sheet = session.activeSheet();
  return (
    <table className="preview">
      <tbody>
        {Array.from({ length: 5 }, (_, row) => (
          <tr key={row}>
            {Array.from({ length: 4 }, (_, col) => (
              <td key={col}>{session.display(sheet.id, row, col)}</td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

async function refreshLinks(session: WorkbookSession) {
  const links = await listLinkedRanges(session.model.id);
  for (const link of links) {
    if (!session.sheet(link.sheetId)) continue;
    link.cells = session.snapshotRange(link.sheetId, link);
    link.sheetName = session.sheet(link.sheetId)?.name ?? link.sheetName;
    link.updatedAt = Date.now();
    await saveLinkedRange(link);
  }
}
