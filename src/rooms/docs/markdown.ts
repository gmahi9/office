import { tableFromRange } from './linkedTable';
import type { DocMark, DocNode } from './model';
import type { LinkedRange } from '../../data/linkedRanges';

export function documentToMarkdown(node: DocNode, tables?: Map<string, LinkedRange>): string {
  return blocks(node.content ?? [], tables).replace(/\n{3,}/g, '\n\n').trim() + '\n';
}

function blocks(nodes: DocNode[], tables?: Map<string, LinkedRange>): string {
  return nodes.map((node) => block(node, tables)).filter(Boolean).join('\n\n');
}

function block(node: DocNode, tables?: Map<string, LinkedRange>): string {
  if (node.type === 'heading') {
    const level = Number(node.attrs?.level ?? 1);
    return `${'#'.repeat(Math.min(6, Math.max(1, level)))} ${inline(node.content)}`;
  }
  if (node.type === 'paragraph') return inline(node.content);
  if (node.type === 'bulletList') {
    return (node.content ?? []).map((item) => `- ${inline(item.content?.[0]?.content)}`).join('\n');
  }
  if (node.type === 'orderedList') {
    return (node.content ?? []).map((item, index) => `${index + 1}. ${inline(item.content?.[0]?.content)}`).join('\n');
  }
  if (node.type === 'taskList') {
    return (node.content ?? []).map((item) => {
      const checked = item.attrs?.checked === true;
      return `- [${checked ? 'x' : ' '}] ${inline(item.content?.[0]?.content)}`;
    }).join('\n');
  }
  if (node.type === 'blockquote') return blocks(node.content ?? [], tables).split('\n').map((line) => `> ${line}`).join('\n');
  if (node.type === 'table') return markdownTable(node);
  if (node.type === 'linkedTable') {
    const id = String(node.attrs?.linkId ?? '');
    const range = tables?.get(id);
    if (!range) return `Linked sheet range (${id})`;
    const model = tableFromRange(range);
    return `${model.caption}\n\n${markdownRows(model.rows)}`;
  }
  if (node.type === 'image') return `![${String(node.attrs?.alt ?? 'image')}](${String(node.attrs?.src ?? '')})`;
  return blocks(node.content ?? [], tables);
}

function markdownTable(node: DocNode): string {
  const rows = (node.content ?? []).map((row) => (row.content ?? []).map((cell) => inline(cell.content?.[0]?.content).replace(/\|/g, '\\|')));
  return markdownRows(rows);
}

function markdownRows(rows: string[][]): string {
  if (!rows.length) return '';
  const width = Math.max(...rows.map((row) => row.length));
  const padded = rows.map((row) => Array.from({ length: width }, (_, index) => row[index] ?? ''));
  const header = padded[0] ?? [];
  const line = header.map(() => '---');
  const body = padded.slice(1);
  return [header, line, ...body].map((row) => `| ${row.join(' | ')} |`).join('\n');
}

function inline(nodes: DocNode[] | undefined): string {
  return (nodes ?? []).map(renderInline).join('');
}

function renderInline(node: DocNode): string {
  if (node.type === 'hardBreak') return '  \n';
  if (node.type !== 'text') return inline(node.content);
  let text = node.text ?? '';
  const marks = node.marks ?? [];
  if (has(marks, 'code')) text = `\`${text}\``;
  if (has(marks, 'bold')) text = `**${text}**`;
  if (has(marks, 'italic')) text = `*${text}*`;
  if (has(marks, 'underline')) text = `<u>${text}</u>`;
  const link = marks.find((mark) => mark.type === 'link');
  if (link) text = `[${text}](${String(link.attrs?.href ?? '')})`;
  return text;
}

function has(marks: DocMark[], type: string): boolean {
  return marks.some((mark) => mark.type === type);
}
