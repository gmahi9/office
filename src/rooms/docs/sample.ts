import type { DocModel } from './model';

export const BRIEF_ID = 'doc_placement_brief';
export const PIPELINE_LINK_ID = 'link_desk_pipeline';

/** A page that reads the placement pipeline already on the desk. */
export function createBriefing(x = 360, y = 64): DocModel {
  const now = Date.now();
  return {
    id: BRIEF_ID,
    title: 'Placement brief',
    createdAt: now,
    updatedAt: now,
    x,
    y,
    comments: [{
      id: 'note_shortlist',
      text: 'Ask the practice lead before the names go out.',
      createdAt: now,
    }],
    content: {
      type: 'doc',
      content: [
        {
          type: 'heading',
          attrs: { level: 1 },
          content: [{ type: 'text', text: 'Placement brief' }],
        },
        {
          type: 'paragraph',
          content: [
            { type: 'text', text: 'The desk is tracking open roles for ' },
            { type: 'text', text: 'AI TEK SOLUTIONS', marks: [{ type: 'bold' }] },
            { type: 'text', text: '. The table is linked to the pipeline sheet, so a change there shows up here.' },
          ],
        },
        {
          type: 'heading',
          attrs: { level: 2 },
          content: [{ type: 'text', text: 'What is moving' }],
        },
        {
          type: 'bulletList',
          content: [
            item('Technology and skilled trades still have openings.'),
            item('Healthcare and maintenance are filled.'),
          ],
        },
        {
          type: 'taskList',
          content: [
            {
              type: 'taskItem',
              attrs: { checked: false },
              content: [{
                type: 'paragraph',
                content: [{
                  type: 'text',
                  text: 'Confirm the Austin analyst shortlist.',
                  marks: [{ type: 'comment', attrs: { id: 'note_shortlist' } }],
                }],
              }],
            },
          ],
        },
        { type: 'linkedTable', attrs: { linkId: PIPELINE_LINK_ID } },
        {
          type: 'paragraph',
          content: [
            { type: 'text', text: 'Rates are on the second sheet. ' },
            {
              type: 'text',
              text: 'Write to the desk',
              marks: [{ type: 'link', attrs: { href: 'https://aiteksolutions.net' } }, { type: 'italic' }],
            },
            { type: 'text', text: ' if a role should be added.' },
          ],
        },
      ],
    },
  };
}

function item(text: string) {
  return {
    type: 'listItem',
    content: [{ type: 'paragraph', content: [{ type: 'text', text }] }],
  };
}
