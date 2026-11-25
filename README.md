# Office

Office is the private, on-device suite for AI TEK SOLUTIONS. Sheets and Docs share one desk. Zoom out to see every open file, and zoom into one to edit it. Slides and Notes are named and not built yet.

Nothing in the app calls a server. Workbooks, documents, linked ranges, and the record of which tools you use stay in IndexedDB in this browser. Opening, editing, and saving a file happens on the device.

Maintainer: gmahi9

## Features

- Virtualized grid, 100,000 rows by 256 columns, with multiple sheets
- Freeze panes, resize and hide rows and columns, merged cells
- Keyboard selection, fill handle, find and replace
- Copy and paste as TSV and HTML, including a paste from Excel or Google Sheets
- Formula engine with cross-sheet references, named ranges, absolute and relative references, and errors
- Number, currency, percent, and date formats, plus type, weight, color, fill, borders, alignment, and wrap
- Conditional rules and color scales, list validation, sort and filter
- Bar, line, pie, and scatter charts from a selection
- Open and save `.xlsx` (values, formulas, basic styles, merges, multiple sheets) and import or export CSV
- Autosave and a recent-files desk
- Tools you use move forward in the toolbar and the command list. Reset that order from Commands
- A linked range can be published from a selection and inserted into a document as a live table
- Docs: headings, lists, checklists, tables, images, links, margin notes, find and replace, word count, and a letter-size page
- Docs open and save `.docx`, export Markdown, and print to PDF
- Source and licences are on the desk: Open source, and Licences

## Keyboard

Sheets:

| Action | Keys |
| --- | --- |
| Move | Arrow keys |
| Extend selection | Shift + arrows |
| Jump to the edge of data | Ctrl + arrows |
| Edit the cell | F2, or type |
| Commit and move | Enter, Shift+Enter, Tab, Shift+Tab |
| Cancel an edit | Escape |
| Copy, cut, paste | Ctrl+C, Ctrl+X, Ctrl+V |
| Undo, redo | Ctrl+Z, Ctrl+Y |
| Bold, italic, underline | Ctrl+B, Ctrl+I, Ctrl+U |
| Find, replace | Ctrl+F, Ctrl+H |
| Fill down, fill right | Ctrl+D, Ctrl+R |
| Select used cells, then the whole sheet | Ctrl+A |
| Commands | Ctrl+K |
| Clear values | Delete |

Ctrl also means Cmd on a Mac keyboard.

Docs, while the page is focused:

| Action | Keys |
| --- | --- |
| Bold, italic, underline | Ctrl+B, Ctrl+I, Ctrl+U |
| Undo, redo | Ctrl+Z, Ctrl+Y |
| Find, replace | Ctrl+F, Ctrl+H |
| Link | Ctrl+K |
| Center text | Ctrl+E |
| Heading 1, 2, or 3 | Ctrl+Alt+1, Ctrl+Alt+2, Ctrl+Alt+3 |
| Save a .docx | Ctrl+S |
| Print / PDF | Ctrl+P |

## Privacy

After the page loads, the app does not send workbook contents anywhere. Files you open are read in the browser. Save and download write a file back to your computer. The desk, autosave, linked ranges, and toolbar memory use IndexedDB on this device only.

## Local development

Requires Node 22.

```bash
npm install
npm run dev
npm test
npm run build
```

The dev server listens on port 43123. `npm run build` typechecks and writes `dist`.

## Cloudflare Pages

- Build command: `npm run build`
- Output directory: `dist`
- `NODE_VERSION`: `22` (also recorded in `.node-version`)

Asset URLs are relative (`base: './'`), so the same build can be served at the zone root or under a path such as `/office/`. `public/_headers` sends a strict content security policy, `frame-ancestors 'none'`, and a long cache for `/assets/*`.

Cell positions use style attributes, so the policy allows `style-src-attr 'unsafe-inline'` and keeps stylesheets and scripts on `'self'`.

## Linked ranges

`src/data/linkedRanges.ts` is the shared store. A sheet publishes a range with its inputs, computed values, and display text. A document reads that record and shows it as a table. It does not need the formula engine.

## Layout

- `src/workspace` is the zoomable desk
- `src/rooms/sheets` is the sheet room
- `src/rooms/docs` is the document room
- `src/rooms/types.ts` lists the other room kinds
- `src/engine` evaluates formulas
- `src/data` is IndexedDB

## Licensing

Our own code is free software under the GNU General Public License, version 3 or any later version (`GPL-3.0-or-later`). The full licence text is in `LICENSE`. There is no warranty.

The suite is a combined work with [HyperFormula](https://hyperformula.handsontable.com/), which is GPL-3.0-only. The engine is initialized with the `gpl-v3` licence key, the key HyperFormula documents for GPL-3.0 projects, and it registers on the order of 400 functions. Because HyperFormula is included, the program you run is distributed as GPL-3.0 open source. You can redistribute it and change it under those terms. Recipients of the combined program do not get a later GPL version from HyperFormula's grant.

Source: https://github.com/gmahi9/office

`THIRD_PARTY_NOTICES.md` lists every bundled dependency, its version, licence, and copyright. `npm run build` regenerates that file and the list shown on the Licences page. The browser build of ExcelJS 4.4.0 (MIT) is vendored in `vendor/` so the suite does not install ExcelJS's Node-only unzip stack. Outfit and Syne are used under the SIL Open Font License.
