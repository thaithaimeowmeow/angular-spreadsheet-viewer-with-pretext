/** Character-encoding detection for text-based formats. */

/** Fallback for non-UTF-8 text: Excel saves "CSV" in the system ANSI code page (Vietnamese Windows = 1258). */
export const LEGACY_ENCODING = 'windows-1258';

export function detectEncoding(head: Uint8Array, declared?: string): string {
  if (head[0] === 0xef && head[1] === 0xbb && head[2] === 0xbf) return 'utf-8';
  if (head[0] === 0xff && head[1] === 0xfe) return 'utf-16le';
  if (head[0] === 0xfe && head[1] === 0xff) return 'utf-16be';
  // BOM-less UTF-16: every other byte is NUL in mostly-ASCII text.
  let evenNul = 0;
  let oddNul = 0;
  const n = Math.min(head.length, 4096);
  for (let i = 0; i < n; i++) if (head[i] === 0) (i & 1 ? oddNul++ : evenNul++);
  if (n > 16 && oddNul > n / 4 && evenNul < n / 50) return 'utf-16le';
  if (n > 16 && evenNul > n / 4 && oddNul < n / 50) return 'utf-16be';

  if (declared) {
    try {
      return new TextDecoder(declared).encoding;
    } catch { /* unknown label, fall through */ }
  }
  try {
    new TextDecoder('utf-8', { fatal: true }).decode(head, { stream: true });
    return 'utf-8';
  } catch {
    return LEGACY_ENCODING;
  }
}

/** True when the bytes look like binary data rather than text. */
export function looksBinary(head: Uint8Array, encoding: string): boolean {
  if (encoding.startsWith('utf-16')) return false;
  const n = Math.min(head.length, 4096);
  let control = 0;
  for (let i = 0; i < n; i++) {
    const b = head[i];
    if (b === 0 || (b < 9) || (b > 13 && b < 32 && b !== 27)) control++;
  }
  return control > n * 0.02;
}

export function decode(bytes: Uint8Array, encoding: string): string {
  return new TextDecoder(encoding).decode(bytes);
}
