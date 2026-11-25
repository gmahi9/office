import Image from '@tiptap/extension-image';
import Link from '@tiptap/extension-link';
import Placeholder from '@tiptap/extension-placeholder';
import { Table, TableCell, TableHeader, TableRow } from '@tiptap/extension-table';
import TaskItem from '@tiptap/extension-task-item';
import TaskList from '@tiptap/extension-task-list';
import TextAlign from '@tiptap/extension-text-align';
import { EditorContent, useEditor, type Editor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import { useEffect, useMemo, useRef, useState } from 'react';
import { Dialog, ProgressButton } from '../../chrome/ui';
import { listLinkedRanges, type LinkedRange } from '../../data/linkedRanges';
import { commandById } from '../../nav/commands';
import { rankCommands } from '../../nav/rank';
import type { UsageStat } from '../../data/usage';
import { DOC_COMMANDS, DOC_TOOL_IDS } from './commands';
import { documentToDocx, docxToHtml } from './docx';
import { CommentMark, LinkedTable } from './extensions';
import { documentToMarkdown } from './markdown';
import { countWords, type DocComment, type DocModel, type DocNode } from './model';

export function DocRoom({
  doc,
  usage,
  reducedMotion,
  onUse,
  onChange,
  onToast,
}: {
  doc: DocModel;
  usage: UsageStat[];
  reducedMotion: boolean;
  onUse: (id: string) => void;
  onChange: (patch: Partial<DocModel>) => void;
  onToast: (message: string, undo?: () => void) => void;
}) {
  const fileRef = useRef<HTMLInputElement>(null);
  const imageRef = useRef<HTMLInputElement>(null);
  const onChangeRef = useRef(onChange);
  onChangeRef.current = onChange;
  const [comments, setComments] = useState<DocComment[]>(doc.comments);
  const commentsRef = useRef(comments);
  commentsRef.current = comments;
  const [dialog, setDialog] = useState<'find' | 'link' | 'comment' | 'ranges' | null>(null);
  const [query, setQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [hits, setHits] = useState<Array<{ from: number; to: number }>>([]);
  const [hitAt, setHitAt] = useState(0);
  const [linkUrl, setLinkUrl] = useState('https://');
  const [commentText, setCommentText] = useState('');
  const [ranges, setRanges] = useState<LinkedRange[]>([]);
  const [words, setWords] = useState(0);

  const editor = useEditor({
    immediatelyRender: true,
    extensions: [
      StarterKit.configure({
        heading: { levels: [1, 2, 3] },
        link: false,
        underline: {},
      }),
      Link.configure({ openOnClick: false, autolink: true }),
      Image.configure({ allowBase64: true }),
      Table.configure({ resizable: false }),
      TableRow,
      TableHeader,
      TableCell,
      TaskList,
      TaskItem.configure({ nested: true }),
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      Placeholder.configure({ placeholder: 'Start the page…' }),
      CommentMark,
      LinkedTable,
    ],
    content: doc.content,
    editorProps: {
      attributes: { 'aria-label': 'Document' },
      handleKeyDown(_view, event) {
        return handleKeys(event);
      },
    },
    onUpdate: ({ editor: current }) => {
      setWords(countWords(current.getText()));
      window.clearTimeout(timer.current);
      timer.current = window.setTimeout(() => {
        onChangeRef.current({
          content: current.getJSON() as DocNode,
          comments: commentsRef.current,
        });
      }, 350);
    },
  });
  const timer = useRef(0);
  const editorRef = useRef<Editor | null>(null);
  editorRef.current = editor;

  useEffect(() => {
    if (editor) setWords(countWords(editor.getText()));
    return () => window.clearTimeout(timer.current);
  }, [editor]);

  const tools = useMemo(() => {
    const defs = DOC_TOOL_IDS.map((id) => DOC_COMMANDS.find((command) => command.id === id)).filter((item) => item != null);
    return rankCommands(defs, usage, Date.now());
  }, [usage]);

  function handleKeys(event: KeyboardEvent): boolean {
    const current = editorRef.current;
    if (!current) return false;
    const mod = event.ctrlKey || event.metaKey;
    if (!mod) return false;
    const key = event.key.toLowerCase();
    if (key === 's') {
      event.preventDefault();
      void saveDocx(async () => undefined);
      return true;
    }
    if (key === 'p') {
      event.preventDefault();
      window.print();
      return true;
    }
    if (key === 'f' || key === 'h') {
      event.preventDefault();
      setDialog('find');
      return true;
    }
    if (key === 'k') {
      event.preventDefault();
      setDialog('link');
      return true;
    }
    if (key === 'e') {
      event.preventDefault();
      current.chain().focus().setTextAlign('center').run();
      onUse('align-center');
      return true;
    }
    if (event.altKey && ['1', '2', '3'].includes(event.key)) {
      event.preventDefault();
      const level = Number(event.key) as 1 | 2 | 3;
      current.chain().focus().toggleHeading({ level }).run();
      onUse(`heading-${level}`);
      return true;
    }
    return false;
  }

  function command(id: string) {
    if (!editor) return;
    onUse(id);
    const chain = editor.chain().focus();
    if (id === 'undo') editor.chain().focus().undo().run();
    if (id === 'redo') editor.chain().focus().redo().run();
    if (id === 'bold') chain.toggleBold().run();
    if (id === 'italic') chain.toggleItalic().run();
    if (id === 'underline') chain.toggleUnderline().run();
    if (id === 'heading-1') chain.toggleHeading({ level: 1 }).run();
    if (id === 'heading-2') chain.toggleHeading({ level: 2 }).run();
    if (id === 'heading-3') chain.toggleHeading({ level: 3 }).run();
    if (id === 'bullet-list') chain.toggleBulletList().run();
    if (id === 'ordered-list') chain.toggleOrderedList().run();
    if (id === 'task-list') chain.toggleTaskList().run();
    if (id === 'align-left') chain.setTextAlign('left').run();
    if (id === 'align-center') chain.setTextAlign('center').run();
    if (id === 'align-right') chain.setTextAlign('right').run();
    if (id === 'insert-table') chain.insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run();
    if (id === 'insert-link') setDialog('link');
    if (id === 'insert-image') imageRef.current?.click();
    if (id === 'comment') setDialog('comment');
    if (id === 'find') setDialog('find');
    if (id === 'insert-linked-range') {
      void listLinkedRanges().then((items) => {
        setRanges(items);
        setDialog('ranges');
      });
    }
  }

  function search() {
    if (!editor || !query) {
      setHits([]);
      return;
    }
    const found = findRanges(editor.state.doc, query, matchCase);
    setHits(found);
    setHitAt(0);
    const first = found[0];
    if (first) editor.chain().focus().setTextSelection(first).scrollIntoView().run();
  }

  function replaceOne() {
    const hit = hits[hitAt];
    if (!editor || !hit) return;
    editor.view.dispatch(editor.state.tr.insertText(replacement, hit.from, hit.to));
    search();
  }

  function replaceAll() {
    if (!editor || !query) return;
    const found = findRanges(editor.state.doc, query, matchCase);
    const tr = editor.state.tr;
    for (const hit of [...found].reverse()) tr.insertText(replacement, hit.from, hit.to);
    editor.view.dispatch(tr);
    onToast(found.length ? `Replaced ${found.length}.` : 'Nothing matched.');
    search();
  }

  async function saveDocx(report: (value: number) => void) {
    if (!editor) return;
    report(0.25);
    const tables: Record<string, string[][]> = {};
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'linkedTable') {
        const id = String(node.attrs.linkId ?? '');
        tables[id] = [];
      }
    });
    const links = await listLinkedRanges();
    for (const id of Object.keys(tables)) {
      const link = links.find((item) => item.id === id);
      if (link) tables[id] = link.cells.map((row) => row.map((cell) => cell.display));
    }
    report(0.6);
    const bytes = await documentToDocx(editor.getJSON() as DocNode, tables);
    report(0.9);
    downloadBytes(`${doc.title}.docx`, bytes, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    report(1);
  }

  async function exportMarkdown(report: (value: number) => void) {
    if (!editor) return;
    report(0.4);
    const links = await listLinkedRanges();
    const map = new Map(links.map((link) => [link.id, link]));
    const text = documentToMarkdown(editor.getJSON() as DocNode, map);
    report(0.85);
    downloadBytes(`${doc.title}.md`, new TextEncoder().encode(text), 'text/markdown');
    report(1);
  }

  async function openDocx(file: File) {
    const html = await docxToHtml(await file.arrayBuffer());
    editor?.commands.setContent(html);
    onToast(`Opened ${file.name}.`);
  }

  const visibleComments = comments.filter((comment) => editor?.getHTML().includes(comment.id));

  return (
    <section className="room doc-room" aria-label={doc.title}>
      <div className="toolbar" role="toolbar" aria-label="Document tools">
        <button type="button" className="tool" onClick={() => command('undo')}>Undo</button>
        <button type="button" className="tool" onClick={() => command('redo')}>Redo</button>
        <span className="tool-gap" />
        <ProgressButton label="Save" doneLabel="Saved" onRun={saveDocx} />
        <ProgressButton label="Markdown" doneLabel="Exported" onRun={exportMarkdown} />
        <ProgressButton label="PDF" doneLabel="Ready" onRun={async (report) => {
          report(0.7);
          await new Promise((resolve) => window.setTimeout(resolve, reducedMotion ? 0 : 180));
          report(1);
          window.print();
        }} />
        <button type="button" className="tool" onClick={() => fileRef.current?.click()}>Open</button>
        <input ref={fileRef} hidden type="file" accept=".docx" onChange={(event) => {
          const file = event.target.files?.[0];
          event.target.value = '';
          if (file) void openDocx(file);
        }} />
        <input ref={imageRef} hidden type="file" accept="image/*" onChange={(event) => {
          const file = event.target.files?.[0];
          event.target.value = '';
          if (!file || !editor) return;
          const reader = new FileReader();
          reader.onload = () => editor.chain().focus().setImage({ src: String(reader.result) }).run();
          reader.readAsDataURL(file);
        }} />
        <span className="tool-gap" />
        {tools.map((tool) => (
          <button
            key={tool.id}
            type="button"
            className={`tool${editor?.isActive(activeName(tool.id)) ? ' is-on' : ''}`}
            title={tool.shortcut}
            onClick={() => command(tool.id)}
          >
            {shortLabel(tool.id)}
          </button>
        ))}
      </div>
      <div className="doc-scroll">
        <div className="doc-layout">
          <article className="page-sheet">
            <EditorContent editor={editor} />
          </article>
          <aside className="doc-margin" aria-label="Comments">
            <h2>Notes</h2>
            {visibleComments.length === 0 ? <p>Select text and add a note. It stays in the margin.</p> : null}
            {visibleComments.map((comment) => (
              <article key={comment.id}>
                <p>{comment.text}</p>
                <button type="button" className="text-btn" onClick={() => {
                  const next = comments.filter((item) => item.id !== comment.id);
                  setComments(next);
                  onChange({ comments: next });
                }}>Remove</button>
              </article>
            ))}
          </aside>
        </div>
      </div>
      <footer className="status">
        <span>{words} {words === 1 ? 'word' : 'words'}</span>
        <span>Saved on this device</span>
      </footer>
      {dialog === 'find' ? (
        <Dialog title="Find and replace" onClose={() => setDialog(null)}>
          <label>Find<input value={query} onChange={(event) => setQuery(event.target.value)} autoFocus /></label>
          <label>Replace<input value={replacement} onChange={(event) => setReplacement(event.target.value)} /></label>
          <label className="check"><input type="checkbox" checked={matchCase} onChange={(event) => setMatchCase(event.target.checked)} /> Match case</label>
          <div className="dialog-actions">
            <button type="button" className="tool" onClick={search}>Find</button>
            <button type="button" className="tool" onClick={() => {
              if (!hits.length || !editor) return;
              const next = (hitAt + 1) % hits.length;
              setHitAt(next);
              const hit = hits[next];
              if (hit) editor.chain().focus().setTextSelection(hit).scrollIntoView().run();
            }}>Next</button>
            <button type="button" className="tool" onClick={replaceOne}>Replace</button>
            <button type="button" className="tool" onClick={replaceAll}>Replace all</button>
          </div>
          <p>{hits.length ? `${hitAt + 1} of ${hits.length}` : 'Search the page.'}</p>
        </Dialog>
      ) : null}
      {dialog === 'link' ? (
        <Dialog title="Link" onClose={() => setDialog(null)}>
          <label>Address<input value={linkUrl} onChange={(event) => setLinkUrl(event.target.value)} autoFocus /></label>
          <div className="dialog-actions">
            <button type="button" className="tool" onClick={() => {
              editor?.chain().focus().extendMarkRange('link').setLink({ href: linkUrl }).run();
              setDialog(null);
            }}>Apply</button>
            <button type="button" className="text-btn" onClick={() => editor?.chain().focus().unsetLink().run()}>Remove</button>
          </div>
        </Dialog>
      ) : null}
      {dialog === 'comment' ? (
        <Dialog title="Margin note" onClose={() => setDialog(null)}>
          <label>Note<textarea value={commentText} onChange={(event) => setCommentText(event.target.value)} rows={3} /></label>
          <div className="dialog-actions">
            <button type="button" className="tool" onClick={() => {
              if (!editor || !commentText.trim()) return;
              const id = crypto.randomUUID();
              editor.chain().focus().setMark('comment', { id }).run();
              const next = [...comments, { id, text: commentText.trim(), createdAt: Date.now() }];
              setComments(next);
              onChange({ comments: next, content: editor.getJSON() as DocNode });
              setCommentText('');
              setDialog(null);
            }}>Add</button>
          </div>
        </Dialog>
      ) : null}
      {dialog === 'ranges' ? (
        <Dialog title="Insert a sheet range" onClose={() => setDialog(null)}>
          {ranges.length === 0 ? <p>Publish a range from a workbook first. In Sheets, select cells and choose Link.</p> : null}
          <ul className="plain-list">
            {ranges.map((range) => (
              <li key={range.id}>
                <button type="button" onClick={() => {
                  editor?.chain().focus().insertContent({ type: 'linkedTable', attrs: { linkId: range.id } }).run();
                  setDialog(null);
                }}>{range.sheetName} {range.a1}</button>
              </li>
            ))}
          </ul>
        </Dialog>
      ) : null}
      <span className="sr-only">{commandById('find')?.label}</span>
    </section>
  );
}

function findRanges(doc: { descendants: (fn: (node: { isText: boolean; text?: string | null }, pos: number) => void) => void }, query: string, matchCase: boolean) {
  const needle = matchCase ? query : query.toLowerCase();
  const hits: Array<{ from: number; to: number }> = [];
  doc.descendants((node, pos) => {
    if (!node.isText || !node.text || !needle) return;
    const text = matchCase ? node.text : node.text.toLowerCase();
    let index = 0;
    while (index <= text.length) {
      const found = text.indexOf(needle, index);
      if (found < 0) break;
      hits.push({ from: pos + found, to: pos + found + needle.length });
      index = found + needle.length;
    }
  });
  return hits;
}

function shortLabel(id: string): string {
  const labels: Record<string, string> = {
    bold: 'Bold',
    italic: 'Italic',
    underline: 'Underline',
    'heading-1': 'Title',
    'heading-2': 'Heading',
    'heading-3': 'Subhead',
    'bullet-list': 'Bullets',
    'ordered-list': 'Numbers',
    'task-list': 'Checklist',
    'align-left': 'Left',
    'align-center': 'Center',
    'align-right': 'Right',
    'insert-link': 'Link',
    'insert-image': 'Image',
    'insert-table': 'Table',
    comment: 'Note',
    'insert-linked-range': 'Sheet',
    find: 'Find',
  };
  return labels[id] ?? id;
}

function activeName(id: string): string {
  if (id === 'bold' || id === 'italic' || id === 'underline') return id;
  if (id === 'heading-1') return 'heading';
  if (id === 'bullet-list') return 'bulletList';
  if (id === 'ordered-list') return 'orderedList';
  if (id === 'task-list') return 'taskList';
  return id;
}

function downloadBytes(filename: string, bytes: Uint8Array, mime: string) {
  const copy = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(copy).set(bytes);
  const url = URL.createObjectURL(new Blob([copy], { type: mime }));
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}
