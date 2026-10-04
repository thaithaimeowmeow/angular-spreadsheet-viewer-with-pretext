/**
 * Spreadsheet worker. The format is detected from the file's bytes (not its extension):
 *   ZIP   → .xlsx/.xlsm/.xltx/.xltm or .ods
 *   OLE2  → legacy .xls (BIFF8/BIFF5)
 *   text  → SpreadsheetML 2003 XML, HTML table (often named .xls), or CSV/TSV/TXT
 *
 * It keeps the current sheet's rows (SheetSession) and serves the UI's queries; the UI never
 * holds the whole sheet.
 */
import { openCsv } from './format-csv';
import { isSpreadsheetMl, openHtml, openSpreadsheetMl } from './format-markup';
import { openOds } from './format-ods';
import { openXls } from './format-xls';
import { openXlsx } from './format-xlsx';
import { type Book, Cancelled, Job, post, setActiveJob } from './job';
import type { ToWorker } from './protocol';
import { SheetSession } from './session';
import { decode, detectEncoding, looksBinary } from './text';
import { ZipArchive } from './zip';

let book: Book | null = null;
let session: SheetSession | null = null;

addEventListener('message', async (e: MessageEvent<ToWorker>) => {
  const msg = e.data;
  if (msg.type !== 'open' && msg.type !== 'sheet') {
    // Queries never cancel work; stale ones (from a previous sheet session) are dropped.
    if (!session || msg.job !== session.job) return;
    switch (msg.type) {
      case 'window': session.window(msg.reqId, msg.start, msg.end); break;
      case 'layout': session.setLayout(msg.version, msg.widths, msg.wrap); break;
      case 'autofit': session.autofit(msg.reqId, msg.col, msg.from); break;
      case 'search': session.setSearch(msg.searchId, msg.query); break;
      case 'copy': session.copy(msg.reqId, msg.r0, msg.r1, msg.c0, msg.c1); break;
    }
    return;
  }

  setActiveJob(msg.job);
  const job = new Job(msg.job);
  session = null; // frees the previous sheet's rows
  try {
    if (msg.type === 'open') {
      book = null;
      job.progress('reading', 0, 1);
      const opened = await detectAndOpen(msg.file, job);
      job.check();
      book = opened;
      post({ type: 'meta', job: job.id, sheets: opened.sheets, format: opened.format });
    } else {
      if (!book) throw new Error('No file is open.');
      if (msg.index < 0 || msg.index >= book.sheets.length) throw new Error(`Sheet ${msg.index + 1} does not exist.`);
      const current = (session = new SheetSession(msg.job, msg.wrap));
      await book.load(msg.index, job, current);
      job.check();
      current.finish();
    }
  } catch (err) {
    if (err instanceof Cancelled) return;
    post({ type: 'error', job: msg.job, message: err instanceof Error ? err.message : String(err) });
  }
});

const baseName = (name: string) => name.replace(/\.[^.]+$/, '') || name;

async function detectAndOpen(file: File, job: Job): Promise<Book> {
  if (file.size === 0) throw new Error('The file is empty.');
  const head = new Uint8Array(await file.slice(0, 64 * 1024).arrayBuffer());

  // ZIP container
  if (head[0] === 0x50 && head[1] === 0x4b && (head[2] === 3 || head[2] === 5 || head[2] === 7)) {
    const zip = await ZipArchive.open(file);
    const mimetype = (await zip.readText('mimetype'))?.trim() ?? '';
    if (mimetype.includes('opendocument.spreadsheet') || (!zip.find('[Content_Types].xml') && zip.find('content.xml'))) {
      if (mimetype && !mimetype.includes('spreadsheet')) throw new Error(`This is an OpenDocument file of type ${mimetype}, not a spreadsheet.`);
      return openOds(zip, job);
    }
    if (zip.find('[Content_Types].xml') || zip.find('_rels/.rels')) {
      if (zip.find('word/document.xml')) throw new Error('This is a Word document, not a spreadsheet.');
      if (zip.find('ppt/presentation.xml')) throw new Error('This is a PowerPoint file, not a spreadsheet.');
      return openXlsx(zip, job);
    }
    throw new Error('This is a ZIP archive, not a spreadsheet.');
  }

  // OLE2 compound file
  if (head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0) {
    return openXls(await file.arrayBuffer(), job);
  }

  if (head[0] === 0x25 && head[1] === 0x50 && head[2] === 0x44 && head[3] === 0x46) {
    throw new Error('This is a PDF file, not a spreadsheet.');
  }

  // Text formats
  const asciiHead = new TextDecoder('latin1').decode(head.subarray(0, 4096));
  const declared = /<meta[^>]+charset\s*=\s*["']?([\w-]+)/i.exec(asciiHead)?.[1]
    ?? /<\?xml[^>]+encoding\s*=\s*["']([\w-]+)/i.exec(asciiHead)?.[1];
  const encoding = detectEncoding(head, declared);
  if (looksBinary(head, encoding)) throw new Error('Unrecognized file format. Supported: .xlsx, .xlsm, .xls, .ods, .csv, .tsv, .txt.');
  const sample = decode(head, encoding);

  if (isSpreadsheetMl(sample)) return openSpreadsheetMl(decode(new Uint8Array(await file.arrayBuffer()), encoding));
  if (/<(html|!doctype\s+html)\b/i.test(sample) || (/<table\b/i.test(sample) && /<t[rd]\b/i.test(sample))) {
    return openHtml(decode(new Uint8Array(await file.arrayBuffer()), encoding), baseName(file.name));
  }
  return openCsv(file, encoding, sample, baseName(file.name));
}
