export interface CommandDef {
  id: string;
  label: string;
  group: 'edit' | 'format' | 'data' | 'view' | 'file';
  shortcut?: string;
  baseOrder: number;
  /** Pinned tools stay at the left of the toolbar. The rest move forward as they are used. */
  pinned?: boolean;
}

export const COMMANDS: CommandDef[] = [
  { id: 'undo', label: 'Undo', group: 'edit', shortcut: 'Ctrl+Z', baseOrder: 0, pinned: true },
  { id: 'redo', label: 'Redo', group: 'edit', shortcut: 'Ctrl+Y', baseOrder: 1, pinned: true },
  { id: 'bold', label: 'Bold', group: 'format', shortcut: 'Ctrl+B', baseOrder: 10 },
  { id: 'italic', label: 'Italic', group: 'format', shortcut: 'Ctrl+I', baseOrder: 11 },
  { id: 'underline', label: 'Underline', group: 'format', shortcut: 'Ctrl+U', baseOrder: 12 },
  { id: 'align-left', label: 'Align left', group: 'format', baseOrder: 20 },
  { id: 'align-center', label: 'Align center', group: 'format', baseOrder: 21 },
  { id: 'align-right', label: 'Align right', group: 'format', baseOrder: 22 },
  { id: 'wrap', label: 'Wrap text', group: 'format', baseOrder: 23 },
  { id: 'format-number', label: 'Number', group: 'format', baseOrder: 30 },
  { id: 'format-currency', label: 'Currency', group: 'format', baseOrder: 31 },
  { id: 'format-percent', label: 'Percent', group: 'format', baseOrder: 32 },
  { id: 'format-date', label: 'Date', group: 'format', baseOrder: 33 },
  { id: 'format-general', label: 'General', group: 'format', baseOrder: 34 },
  { id: 'fill-color', label: 'Fill color', group: 'format', baseOrder: 40 },
  { id: 'font-color', label: 'Text color', group: 'format', baseOrder: 41 },
  { id: 'border-all', label: 'All borders', group: 'format', baseOrder: 50 },
  { id: 'border-outer', label: 'Outer borders', group: 'format', baseOrder: 51 },
  { id: 'border-bottom', label: 'Bottom border', group: 'format', baseOrder: 52 },
  { id: 'border-none', label: 'Clear borders', group: 'format', baseOrder: 53 },
  { id: 'merge', label: 'Merge cells', group: 'format', baseOrder: 60 },
  { id: 'freeze', label: 'Freeze panes', group: 'view', baseOrder: 61 },
  { id: 'sort-asc', label: 'Sort ascending', group: 'data', baseOrder: 70 },
  { id: 'sort-desc', label: 'Sort descending', group: 'data', baseOrder: 71 },
  { id: 'filter', label: 'Filter', group: 'data', baseOrder: 72 },
  { id: 'chart', label: 'Chart', group: 'data', baseOrder: 73 },
  { id: 'conditional', label: 'Conditional format', group: 'format', baseOrder: 80 },
  { id: 'validation', label: 'Data validation', group: 'data', baseOrder: 81 },
  { id: 'name-range', label: 'Name range', group: 'data', baseOrder: 82 },
  { id: 'link-range', label: 'Link range', group: 'data', baseOrder: 83 },
  { id: 'find', label: 'Find', group: 'edit', shortcut: 'Ctrl+F', baseOrder: 90 },
  { id: 'replace', label: 'Replace', group: 'edit', shortcut: 'Ctrl+H', baseOrder: 91 },
  { id: 'insert-row', label: 'Insert row', group: 'edit', baseOrder: 100 },
  { id: 'insert-col', label: 'Insert column', group: 'edit', baseOrder: 101 },
  { id: 'delete-row', label: 'Delete row', group: 'edit', baseOrder: 102 },
  { id: 'delete-col', label: 'Delete column', group: 'edit', baseOrder: 103 },
  { id: 'hide-row', label: 'Hide row', group: 'view', baseOrder: 110 },
  { id: 'hide-col', label: 'Hide column', group: 'view', baseOrder: 111 },
  { id: 'unhide', label: 'Unhide rows and columns', group: 'view', baseOrder: 112 },
  { id: 'fill-down', label: 'Fill down', group: 'edit', shortcut: 'Ctrl+D', baseOrder: 120 },
  { id: 'fill-right', label: 'Fill right', group: 'edit', shortcut: 'Ctrl+R', baseOrder: 121 },
];

export const TOOLBAR_IDS = [
  'bold',
  'italic',
  'underline',
  'align-left',
  'align-center',
  'align-right',
  'wrap',
  'format-currency',
  'format-percent',
  'format-number',
  'format-date',
  'fill-color',
  'font-color',
  'border-all',
  'merge',
  'freeze',
  'sort-asc',
  'sort-desc',
  'filter',
  'chart',
  'conditional',
  'validation',
  'link-range',
] as const;

export function commandById(id: string): CommandDef | undefined {
  return COMMANDS.find((command) => command.id === id);
}
