export interface DocMark {
  type: string;
  attrs?: Record<string, unknown>;
}

export interface DocNode {
  type: string;
  text?: string;
  attrs?: Record<string, unknown>;
  marks?: DocMark[];
  content?: DocNode[];
}

export interface DocComment {
  id: string;
  text: string;
  createdAt: number;
}

export interface DocModel {
  id: string;
  title: string;
  createdAt: number;
  updatedAt: number;
  x: number;
  y: number;
  content: DocNode;
  comments: DocComment[];
}

export function emptyDocument(title = 'Untitled', x = 80, y = 80): DocModel {
  const now = Date.now();
  return {
    id: crypto.randomUUID(),
    title,
    createdAt: now,
    updatedAt: now,
    x,
    y,
    comments: [],
    content: { type: 'doc', content: [{ type: 'paragraph' }] },
  };
}

export function documentText(node: DocNode | undefined): string {
  if (!node) return '';
  if (node.text) return node.text;
  const gap = node.type === 'doc' || node.type === 'bulletList' || node.type === 'orderedList' || node.type === 'taskList' ? '\n' : '';
  return (node.content ?? []).map(documentText).join(node.type === 'paragraph' || node.type === 'heading' || node.type === 'listItem' || node.type === 'taskItem' ? '\n' : gap);
}

export function countWords(text: string): number {
  const trimmed = text.trim();
  if (!trimmed) return 0;
  return trimmed.split(/\s+/).filter(Boolean).length;
}

export function previewText(doc: DocModel): string {
  return documentText(doc.content).replace(/\s+/g, ' ').trim().slice(0, 140);
}
