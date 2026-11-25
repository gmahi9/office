import { Mark, Node, mergeAttributes } from '@tiptap/core';
import { ReactNodeViewRenderer } from '@tiptap/react';
import { LinkedTableView } from './LinkedTableView';

export const CommentMark = Mark.create({
  name: 'comment',
  inclusive: false,
  addAttributes() {
    return { id: { default: null } };
  },
  parseHTML() {
    return [{ tag: 'span[data-comment]', getAttrs: (element) => ({ id: (element as HTMLElement).getAttribute('data-comment') }) }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', mergeAttributes(HTMLAttributes, { 'data-comment': HTMLAttributes.id, class: 'doc-comment' }), 0];
  },
});

export const LinkedTable = Node.create({
  name: 'linkedTable',
  group: 'block',
  atom: true,
  draggable: true,
  addAttributes() {
    return { linkId: { default: '' } };
  },
  parseHTML() {
    return [{
      tag: 'div[data-linked-table]',
      getAttrs: (element) => ({ linkId: (element as HTMLElement).getAttribute('data-linked-table') }),
    }];
  },
  renderHTML({ HTMLAttributes }) {
    return ['div', mergeAttributes(HTMLAttributes, { 'data-linked-table': HTMLAttributes.linkId })];
  },
  addNodeView() {
    return ReactNodeViewRenderer(LinkedTableView);
  },
});
