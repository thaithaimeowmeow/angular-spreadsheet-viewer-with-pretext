# Angular spreadsheet viewer with pretext

A fast, lightweight spreadsheet previewer built with Angular only (no grid or spreadsheet
libraries). Wrapped cell text is measured with [`@chenglou/pretext`](https://github.com/chenglou/pretext)
instead of the DOM.

## Formats

Detected from the file's bytes, not its extension:

| Format | Notes |
|---|---|
| `.xlsx` `.xlsm` `.xltx` `.xltm` | Streaming unzip (native `DecompressionStream`), namespace-prefixed XML, ZIP64 |
| `.xls` | Excel 97–2003 (BIFF8) and Excel 5/95 (BIFF5) |
| `.ods` | OpenDocument spreadsheets |
| `.csv` `.tsv` `.txt` | Delimiter and encoding detection (UTF-8/16, falls back to windows-1258) |
| HTML tables, Excel 2003 XML | Often saved with an `.xls` extension |

## Features

- Virtualized grid: only the cells on screen exist in the DOM; DOM elements are reused while scrolling.
- Wrapped text with exact row heights for visible rows (pretext), estimates for the rest.
- Search (Ctrl+F): accent-insensitive for Vietnamese, Enter / Shift+Enter to jump, highlighted matches.
- Range selection (drag, Shift+click, row/column headers, Ctrl+A) and Ctrl+C as tab-separated text.
- Resize columns and rows; double-click an edge to auto-fit; double-click a cell (or F2) to read its full text.
- Multiple sheets; very large sheets (> 500k rows) start with wrapping off.

## How it stays light

- **The worker owns the data.** Parsing happens in a Web Worker that keeps the current sheet's rows.
  The page only holds row heights (a Fenwick tree) and a 64-block LRU of rows near the viewport,
  fetched on demand.
- **No pretext off screen.** The worker estimates heights from per-character widths; the page runs
  pretext only on visible rows, where exactness matters. Rows that can't fit in 8 lines are capped
  without measuring.
- **Search and copy run in the worker**, so the page never scans or copies the whole sheet.
- Repeated short values are stored once; long values are copied out of their parse buffers so the
  raw XML/CSV can be garbage-collected.

## Development

```bash
npm install
npx ng serve
```

Open http://localhost:4200 and drop a spreadsheet on the page.

Generate a large test workbook (written to `samples/`, which is git-ignored):

```bash
node scripts/make-sample.mjs 1000000 samples/big.xlsx
```

Run the parser against files or folders (`--ref` compares with SheetJS, used only as a test oracle):

```bash
node scripts/test-parser.mjs path/to/files --ref path/to/node_modules/xlsx
```
