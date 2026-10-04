// Generates a large .xlsx for testing: node scripts/make-sample.mjs [rows] [out]   (e.g. 1000000 samples/big.xlsx)
import { createWriteStream } from 'node:fs';
import { crc32, deflateRawSync } from 'node:zlib';

const ROWS = Number(process.argv[2] ?? 200_000);
const OUT = process.argv[3] ?? `sample-${ROWS}.xlsx`;

const words = 'lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua Tiếng Việt 日本語のテキスト émoji 🚀 naïve'.split(' ');
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
let seed = 42;
const rand = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
const sentence = (n) => Array.from({ length: n }, () => words[Math.floor(rand() * words.length)]).join(' ');
const col = (c) => String.fromCharCode(65 + c);

const shared = ['North', 'South', 'East', 'West', 'Open', 'Closed', 'Pending'];
const headers = ['ID', 'Region', 'Status', 'Amount', 'Ratio', 'Date', 'Notes', 'Comment'];

function sheetXml(rows) {
  const parts = [`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:H${rows + 1}"/><sheetData>`];
  parts.push(`<row r="1">${headers.map((h, c) => `<c r="${col(c)}1" t="inlineStr"><is><t>${h}</t></is></c>`).join('')}</row>`);
  for (let i = 0; i < rows; i++) {
    const r = i + 2;
    const notes = sentence(2 + Math.floor(rand() * (i % 7 === 0 ? 40 : 6)));
    const comment = i % 11 === 0 ? `Line one ${i}\nLine two\nLine three` : '';
    parts.push(
      `<row r="${r}">` +
      `<c r="A${r}"><v>${i + 1}</v></c>` +
      `<c r="B${r}" t="s"><v>${i % 4}</v></c>` +
      `<c r="C${r}" t="s"><v>${4 + (i % 3)}</v></c>` +
      `<c r="D${r}"><v>${(rand() * 100000).toFixed(2)}</v></c>` +
      `<c r="E${r}" s="2"><v>${rand().toFixed(4)}</v></c>` +
      `<c r="F${r}" s="1"><v>${45000 + (i % 3000)}</v></c>` +
      `<c r="G${r}" t="inlineStr"><is><t>${esc(notes)}</t></is></c>` +
      (comment ? `<c r="H${r}" t="inlineStr"><is><t xml:space="preserve">${esc(comment)}</t></is></c>` : '') +
      `</row>`,
    );
  }
  parts.push('</sheetData></worksheet>');
  return parts.join('');
}

const files = {
  '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  '_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Big data" sheetId="1" r:id="rId1"/><sheet name="Small &amp; sparse" sheetId="2" r:id="rId2"/><sheet name="Empty" sheetId="3" r:id="rId3"/></sheets></workbook>`,
  'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${shared.map((s, i) => i === 0 ? `<si><r><t>No</t></r><r><t>rth</t></r><rPh><t>XX</t></rPh></si>` : `<si><t>${s}</t></si>`).join('')}</sst>`,
  'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/></numFmts><cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="164" applyNumberFormat="1"/><xf numFmtId="10" applyNumberFormat="1"/></cellXfs></styleSheet>`,
  'xl/worksheets/sheet1.xml': sheetXml(ROWS),
  'xl/worksheets/sheet2.xml': `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2" t="inlineStr"><is><t>Sparse cell &amp; entity &lt;ok&gt; _x0041_</t></is></c></row><row r="5"><c r="D5" t="b"><v>1</v></c><c r="F5" t="e"><v>#DIV/0!</v></c></row><row r="40"><c r="AB40"><v>3.14159265358979</v></c></row></sheetData></worksheet>`,
  'xl/worksheets/sheet3.xml': `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
};

const out = createWriteStream(OUT);
const central = [];
let offset = 0;
const write = (buf) => { out.write(buf); offset += buf.length; };

for (const [name, text] of Object.entries(files)) {
  const raw = Buffer.from(text, 'utf8');
  const data = deflateRawSync(raw, { level: 6 });
  const nameBuf = Buffer.from(name);
  const crc = crc32(raw);
  const local = Buffer.alloc(30);
  local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x0800, 6);
  local.writeUInt16LE(8, 8); local.writeUInt32LE(crc, 14); local.writeUInt32LE(data.length, 18);
  local.writeUInt32LE(raw.length, 22); local.writeUInt16LE(nameBuf.length, 26);
  central.push({ nameBuf, crc, comp: data.length, size: raw.length, offset });
  write(local); write(nameBuf); write(data);
  console.log(`${name}: ${(raw.length / 1e6).toFixed(1)} MB → ${(data.length / 1e6).toFixed(1)} MB`);
}

const cdStart = offset;
for (const e of central) {
  const h = Buffer.alloc(46);
  h.writeUInt32LE(0x02014b50, 0); h.writeUInt16LE(20, 4); h.writeUInt16LE(20, 6); h.writeUInt16LE(0x0800, 8);
  h.writeUInt16LE(8, 10); h.writeUInt32LE(e.crc, 16); h.writeUInt32LE(e.comp, 20); h.writeUInt32LE(e.size, 24);
  h.writeUInt16LE(e.nameBuf.length, 28); h.writeUInt32LE(e.offset, 42);
  write(h); write(e.nameBuf);
}
const end = Buffer.alloc(22);
end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(central.length, 8); end.writeUInt16LE(central.length, 10);
end.writeUInt32LE(offset - cdStart, 12); end.writeUInt32LE(cdStart, 16);
write(end);
out.end(() => console.log(`wrote ${OUT} (${(offset / 1e6).toFixed(1)} MB)`));
