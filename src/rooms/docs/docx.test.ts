import { describe, expect, it } from 'vitest';
import { documentToDocx, docxToHtml, restoreLinkedMarkers } from './docx';
import { documentToMarkdown } from './markdown';
import type { DocNode } from './model';
import type { LinkedRange } from '../../data/linkedRanges';

const page: DocNode = {
  type: 'doc',
  content: [
    { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Placement brief' }] },
    {
      type: 'paragraph',
      content: [
        { type: 'text', text: 'Open roles for ' },
        { type: 'text', text: 'the desk', marks: [{ type: 'bold' }] },
        { type: 'text', text: ' stay ' },
        { type: 'text', text: 'in motion', marks: [{ type: 'italic' }] },
        { type: 'text', text: ' and ' },
        { type: 'text', text: 'underlined', marks: [{ type: 'underline' }] },
        { type: 'text', text: '. ' },
        { type: 'text', text: 'Write to the desk', marks: [{ type: 'link', attrs: { href: 'https://aiteksolutions.net' } }] },
      ],
    },
    {
      type: 'bulletList',
      content: [
        { type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Austin analyst' }] }] },
      ],
    },
    {
      type: 'table',
      content: [
        {
          type: 'tableRow',
          content: [
            { type: 'tableHeader', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Role' }] }] },
            { type: 'tableHeader', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Openings' }] }] },
          ],
        },
        {
          type: 'tableRow',
          content: [
            { type: 'tableCell', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Analyst' }] }] },
            { type: 'tableCell', content: [{ type: 'paragraph', content: [{ type: 'text', text: '4' }] }] },
          ],
        },
      ],
    },
    { type: 'linkedTable', attrs: { linkId: 'link_desk_pipeline' } },
  ],
};

describe('docx round trip', () => {
  it('keeps headings, marks, lists, tables, links, and a linked-range marker', async () => {
    const bytes = await documentToDocx(page, { link_desk_pipeline: [['Role', 'Openings'], ['Analyst', '4']] });
    const html = await docxToHtml(bytes);
    expect(html.toLowerCase()).toContain('placement brief');
    expect(html).toMatch(/<(strong|b)>the desk<\/(strong|b)>/i);
    expect(html).toMatch(/<(em|i)>in motion<\/(em|i)>/i);
    expect(html.toLowerCase()).toMatch(/<u>underlined<\/u>|text-decoration:\s*underline/);
    expect(html.toLowerCase()).toContain('austin analyst');
    expect(html.toLowerCase()).toContain('<table');
    expect(html).toContain('https://aiteksolutions.net');
    expect(html).toContain('data-linked-table="link_desk_pipeline"');
    expect(restoreLinkedMarkers('<p>Linked from Sheets (abc)</p><table><tr><td>1</td></tr></table>')).toBe(
      '<div data-linked-table="abc"></div>',
    );
  });
});

describe('markdown', () => {
  it('writes a heading, a list, and a linked table', () => {
    const range: LinkedRange = {
      id: 'link_desk_pipeline',
      workbookId: 'book',
      sheetId: 'sheet',
      sheetName: 'Pipeline',
      a1: 'A4:B5',
      r1: 3,
      c1: 0,
      r2: 4,
      c2: 1,
      cells: [
        [{ input: 'Role', value: 'Role', display: 'Role' }, { input: 'Openings', value: 'Openings', display: 'Openings' }],
        [{ input: 'Analyst', value: 'Analyst', display: 'Analyst' }, { input: '4', value: 4, display: '4' }],
      ],
      createdAt: 1,
      updatedAt: 1,
    };
    const text = documentToMarkdown(page, new Map([[range.id, range]]));
    expect(text).toContain('# Placement brief');
    expect(text).toContain('- Austin analyst');
    expect(text).toContain('| Role | Openings |');
    expect(text).toContain('| Analyst | 4 |');
  });
});
