import type { CommandDef } from '../../nav/commands';

export const DOC_COMMANDS: CommandDef[] = [
  { id: 'undo', label: 'Undo', group: 'edit', shortcut: 'Ctrl+Z', baseOrder: 0, pinned: true },
  { id: 'redo', label: 'Redo', group: 'edit', shortcut: 'Ctrl+Y', baseOrder: 1, pinned: true },
  { id: 'bold', label: 'Bold', group: 'format', shortcut: 'Ctrl+B', baseOrder: 10 },
  { id: 'italic', label: 'Italic', group: 'format', shortcut: 'Ctrl+I', baseOrder: 11 },
  { id: 'underline', label: 'Underline', group: 'format', shortcut: 'Ctrl+U', baseOrder: 12 },
  { id: 'heading-1', label: 'Heading 1', group: 'format', shortcut: 'Ctrl+Alt+1', baseOrder: 20 },
  { id: 'heading-2', label: 'Heading 2', group: 'format', shortcut: 'Ctrl+Alt+2', baseOrder: 21 },
  { id: 'heading-3', label: 'Heading 3', group: 'format', shortcut: 'Ctrl+Alt+3', baseOrder: 22 },
  { id: 'bullet-list', label: 'Bullets', group: 'format', baseOrder: 30 },
  { id: 'ordered-list', label: 'Numbered', group: 'format', baseOrder: 31 },
  { id: 'task-list', label: 'Checklist', group: 'format', baseOrder: 32 },
  { id: 'align-left', label: 'Align left', group: 'format', baseOrder: 40 },
  { id: 'align-center', label: 'Align center', group: 'format', shortcut: 'Ctrl+E', baseOrder: 41 },
  { id: 'align-right', label: 'Align right', group: 'format', baseOrder: 42 },
  { id: 'insert-link', label: 'Link', group: 'edit', shortcut: 'Ctrl+K', baseOrder: 50 },
  { id: 'insert-image', label: 'Image', group: 'edit', baseOrder: 51 },
  { id: 'insert-table', label: 'Table', group: 'edit', baseOrder: 52 },
  { id: 'comment', label: 'Comment', group: 'edit', baseOrder: 60 },
  { id: 'insert-linked-range', label: 'Sheet range', group: 'data', baseOrder: 70 },
  { id: 'find', label: 'Find', group: 'edit', shortcut: 'Ctrl+F', baseOrder: 80 },
];

export const DOC_TOOL_IDS = [
  'bold', 'italic', 'underline', 'heading-1', 'heading-2', 'heading-3',
  'bullet-list', 'ordered-list', 'task-list', 'align-left', 'align-center', 'align-right',
  'insert-link', 'insert-image', 'insert-table', 'comment', 'insert-linked-range', 'find',
];
