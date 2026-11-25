export type RoomKind = 'sheet' | 'doc' | 'slide' | 'note';

export interface RoomKindInfo {
  kind: RoomKind;
  label: string;
  available: boolean;
  blurb: string;
}

/** The desk knows every room kind. Sheets and Docs are built. */
export const ROOM_KINDS: RoomKindInfo[] = [
  { kind: 'sheet', label: 'Sheets', available: true, blurb: 'Workbooks on this device.' },
  { kind: 'doc', label: 'Docs', available: true, blurb: 'Pages on this device.' },
  { kind: 'slide', label: 'Slides', available: false, blurb: 'A frame will be able to link a range.' },
  { kind: 'note', label: 'Notes', available: false, blurb: 'Short notes will sit beside the files.' },
];
