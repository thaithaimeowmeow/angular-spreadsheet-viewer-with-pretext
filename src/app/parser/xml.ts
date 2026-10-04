/** Minimal, streaming-friendly XML helpers (regex based, namespace-prefix tolerant). */
import type { Job } from './job';

// Element names may carry any namespace prefix: <row>, <x:row>, <ns0:row> are all valid
// (OpenXML SDK / .NET exporters write prefixed XML), so every tag pattern allows one.
export const P = '(?:[\\w.-]+:)?';

/** Matches <tag attrs/> or <tag attrs>inner</tag>; group 1 = attrs, group 2 = inner. */
export function element(tag: string, flags = 'g'): RegExp {
  return new RegExp(`<${P}${tag}(?=[\\s/>])([^>]*?)(?:/>|>([\\s\\S]*?)</${P}${tag}\\s*>)`, flags);
}

/** End index of the last complete closing tag of `tag` (any prefix) in buf, or -1. */
export function lastCloseEnd(buf: string, tag: string): number {
  const needle = tag + '>';
  for (let i = buf.lastIndexOf(needle); i > 0; i = buf.lastIndexOf(needle, i - 1)) {
    let j = i - 1;
    if (buf[j] === ':') {
      j--;
      while (j >= 0 && /[\w.-]/.test(buf[j])) j--;
    }
    if (buf[j] === '/' && buf[j - 1] === '<') return i + needle.length;
  }
  return -1;
}

const ENTITY = /&(?:(amp|lt|gt|quot|apos)|#(\d+)|#[xX]([0-9a-fA-F]+));/g;
const NAMED: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
const OOXML_ESC = /_x([0-9A-Fa-f]{4})_/g;

function codePoint(dec: string | undefined, hex: string | undefined): string {
  const cp = dec ? +dec : parseInt(hex!, 16);
  return cp > 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : '';
}

/**
 * Returns a copy of `s` that does not keep its parent string alive. V8 represents substrings of
 * 13+ chars (regex captures, slice) as views into the parent, so a stored cell value would pin
 * the whole XML/CSV chunk it came from; concatenating then slicing forces a flat copy.
 * [V8 behavior, not a language guarantee]
 */
export function own(s: string): string {
  return s.length < 13 ? s : (' ' + s).slice(1);
}

/** Normalizes CR / CRLF to LF: the DOM renders a lone CR as a space, pretext must agree. */
export function normalizeNewlines(s: string): string {
  return s.indexOf('\r') === -1 ? s : s.replace(/\r\n?/g, '\n');
}

export function unescapeXml(s: string, ooxml = true): string {
  if (s.indexOf('&') !== -1) {
    s = s.replace(ENTITY, (_, named: string, dec: string, hex: string) => (named ? NAMED[named] : codePoint(dec, hex)));
  }
  if (ooxml && s.indexOf('_x') !== -1) {
    s = s.replace(OOXML_ESC, (_, hex: string) => String.fromCharCode(parseInt(hex, 16)));
  }
  return normalizeNewlines(s);
}

/** Attribute value by local name, ignoring any prefix (`r:id`, `ss:Name`, `table:name`...). */
export function attr(attrs: string, local: string, prefixed = false): string | undefined {
  const pre = prefixed ? P : '';
  const m = new RegExp(`(?:^|\\s)${pre}${local}\\s*=\\s*(["'])(.*?)\\1`).exec(attrs);
  return m ? unescapeXml(m[2], false) : undefined;
}

/**
 * Streams text, handing `process` chunks that end right after the last complete closing tag of
 * any of `cutTags`, so no element of interest is ever split. Keeps only the unprocessed tail.
 * `process` may return true to stop early.
 */
export async function scanStream(
  stream: ReadableStream<string>,
  cutTags: string[],
  job: Job,
  process: (chunk: string) => boolean | void,
): Promise<void> {
  const reader = stream.getReader();
  let buf = '';
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      await job.tick();
      buf += value;
      let end = -1;
      for (const tag of cutTags) end = Math.max(end, lastCloseEnd(buf, tag));
      if (end === -1) continue;
      if (process(buf.slice(0, end))) {
        reader.cancel().catch(() => {});
        return;
      }
      buf = buf.slice(end);
    }
    process(buf);
  } catch (err) {
    reader.cancel().catch(() => {});
    throw err;
  }
}

/** Streams every `<tag>` element (any prefix) to `onElement`. */
export function scanElements(
  stream: ReadableStream<string>,
  tag: string,
  job: Job,
  onElement: (attrs: string, inner: string | undefined) => void,
): Promise<void> {
  const re = element(tag);
  return scanStream(stream, [tag], job, (chunk) => {
    re.lastIndex = 0;
    for (let m = re.exec(chunk); m; m = re.exec(chunk)) onElement(m[1], m[2]);
  });
}
