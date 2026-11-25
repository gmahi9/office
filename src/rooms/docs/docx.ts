import {
  AlignmentType,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  UnderlineType,
  WidthType,
  type ParagraphChild,
} from 'docx';
import mammoth from 'mammoth';
import type { DocNode } from './model';

const HEADING: Record<number, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
};

export async function documentToDocx(
  content: DocNode,
  tables: Record<string, string[][]> = {},
): Promise<Uint8Array> {
  const children = blocks(content.content ?? [], tables);
  const doc = new Document({
    title: 'Office document',
    creator: 'AI TEK SOLUTIONS',
    numbering: {
      config: [
        {
          reference: 'office-bullets',
          levels: [{
            level: 0,
            format: LevelFormat.BULLET,
            text: '•',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          }],
        },
        {
          reference: 'office-numbers',
          levels: [{
            level: 0,
            format: LevelFormat.DECIMAL,
            text: '%1.',
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          }],
        },
      ],
    },
    sections: [{ children }],
  });
  const out = await Packer.toBuffer(doc);
  return new Uint8Array(out);
}

export async function docxToHtml(data: ArrayBuffer | Uint8Array): Promise<string> {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const copy = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(copy).set(bytes);
  const result = await mammoth.convertToHtml({ buffer: bytes, arrayBuffer: copy } as never, {
    styleMap: ['u => u'],
    convertImage: mammoth.images.imgElement(async (image) => {
      const base64 = await image.read('base64');
      return { src: `data:${image.contentType};base64,${base64}` };
    }),
  });
  return restoreLinkedMarkers(result.value);
}

/** Turn the export marker and its static table back into a live linked table. */
export function restoreLinkedMarkers(html: string): string {
  return html.replace(
    /<p>Linked from Sheets \(([^)<]+)\)<\/p>\s*<table[\s\S]*?<\/table>/gi,
    (_all, id: string) => `<div data-linked-table="${id}"></div>`,
  );
}

function blocks(nodes: DocNode[], tables: Record<string, string[][]>): Array<Paragraph | Table> {
  const out: Array<Paragraph | Table> = [];
  for (const node of nodes) {
    if (node.type === 'bulletList' || node.type === 'orderedList' || node.type === 'taskList') {
      const reference = node.type === 'orderedList' ? 'office-numbers' : node.type === 'bulletList' ? 'office-bullets' : '';
      for (const item of node.content ?? []) {
        const checked = item.attrs?.checked === true;
        const prefix = node.type === 'taskList' ? (checked ? '☑ ' : '☐ ') : '';
        out.push(new Paragraph({
          numbering: reference ? { reference, level: 0 } : undefined,
          children: runs(item.content?.[0]?.content, prefix),
        }));
      }
      continue;
    }
    if (node.type === 'table') {
      out.push(docxTable((node.content ?? []).map((row) => (row.content ?? []).map((cell) => inlineText(cell.content)))));
      continue;
    }
    if (node.type === 'linkedTable') {
      const id = String(node.attrs?.linkId ?? '');
      out.push(new Paragraph({ children: [new TextRun(`Linked from Sheets (${id})`)] }));
      const rows = tables[id];
      if (rows?.length) out.push(docxTable(rows.map((row) => row.map((cell) => cell))));
      continue;
    }
    if (node.type === 'image') {
      const image = imageRun(String(node.attrs?.src ?? ''));
      out.push(new Paragraph({ children: image ? [image] : [new TextRun('')] }));
      continue;
    }
    if (node.type === 'blockquote') {
      out.push(...blocks(node.content ?? [], tables));
      continue;
    }
    const level = Number(node.attrs?.level ?? 0);
    out.push(new Paragraph({
      heading: node.type === 'heading' ? HEADING[level] ?? HeadingLevel.HEADING_1 : undefined,
      alignment: alignment(node),
      children: runs(node.content),
    }));
  }
  return out;
}

function docxTable(rows: string[][]): Table {
  const width = Math.max(1, ...rows.map((row) => row.length));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows.map((row) => new TableRow({
      children: Array.from({ length: width }, (_, index) => new TableCell({
        width: { size: 100 / width, type: WidthType.PERCENTAGE },
        children: [new Paragraph({ children: [new TextRun(row[index] ?? '')] })],
      })),
    })),
  });
}

function runs(nodes: DocNode[] | undefined, prefix = ''): ParagraphChild[] {
  const children: ParagraphChild[] = [];
  let prefixed = false;
  for (const node of nodes ?? []) {
    if (node.type === 'hardBreak') {
      children.push(new TextRun({ break: 1 }));
      continue;
    }
    if (node.type !== 'text') {
      children.push(...runs(node.content));
      continue;
    }
    const text = `${prefixed ? '' : prefix}${node.text ?? ''}`;
    prefixed = true;
    const marks = node.marks ?? [];
    const run = new TextRun({
      text,
      bold: marks.some((mark) => mark.type === 'bold') || undefined,
      italics: marks.some((mark) => mark.type === 'italic') || undefined,
      underline: marks.some((mark) => mark.type === 'underline') ? { type: UnderlineType.SINGLE } : undefined,
    });
    const link = marks.find((mark) => mark.type === 'link');
    if (link) {
      children.push(new ExternalHyperlink({ link: String(link.attrs?.href ?? ''), children: [run] }));
    } else {
      children.push(run);
    }
  }
  if (!prefixed && prefix) children.unshift(new TextRun(prefix));
  if (!children.length) children.push(new TextRun(''));
  return children;
}

function inlineText(nodes: DocNode[] | undefined): string {
  return (nodes ?? []).map((node) => node.text ?? inlineText(node.content)).join('');
}

function alignment(node: DocNode) {
  const value = node.attrs?.textAlign;
  if (value === 'center') return AlignmentType.CENTER;
  if (value === 'right') return AlignmentType.RIGHT;
  if (value === 'justify') return AlignmentType.BOTH;
  return undefined;
}

function imageRun(src: string): ImageRun | null {
  const match = /^data:image\/(png|jpeg|jpg|gif|bmp);base64,(.+)$/i.exec(src);
  if (!match) return null;
  const raw = (match[1] ?? 'png').toLowerCase();
  const type = raw === 'jpeg' ? 'jpg' : raw;
  const bytes = Uint8Array.from(atob(match[2] ?? ''), (char) => char.charCodeAt(0));
  return new ImageRun({
    type: type as 'png' | 'jpg' | 'gif' | 'bmp',
    data: bytes,
    transformation: { width: 480, height: 300 },
  });
}
