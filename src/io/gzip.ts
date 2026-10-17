/**
 * Gzip (RFC 1952) and raw DEFLATE decompression for imported files: gzipped Gnumeric and
 * KSpread documents, Apple Numbers '09 index.xml.gz, and any *.gz file around a spreadsheet.
 *
 * Uses the platform's DecompressionStream when available and falls back to the project's own
 * DEFLATE decoder (recover/inflate.ts), so it works in every runtime the app ships in.
 */

export const isGzip = (u8: Uint8Array): boolean => u8.length > 18 && u8[0] === 0x1f && u8[1] === 0x8b && u8[2] === 8;

/** Offset of the DEFLATE data in a gzip member (after the optional header fields). */
function gzipDataOffset(u8: Uint8Array, p: number): number {
  const flg = u8[p + 3];
  let o = p + 10;
  if (flg & 4) o += 2 + (u8[o] | (u8[o + 1] << 8));
  if (flg & 8) { while (o < u8.length && u8[o]) o++; o++; }
  if (flg & 16) { while (o < u8.length && u8[o]) o++; o++; }
  if (flg & 2) o += 2;
  return o;
}

async function viaStream(u8: Uint8Array, format: 'gzip' | 'deflate-raw'): Promise<Uint8Array | null> {
  const DS = (globalThis as any).DecompressionStream;
  if (typeof DS !== 'function' || typeof Response === 'undefined') return null;
  try {
    const stream = new Blob([u8 as BlobPart]).stream().pipeThrough(new DS(format));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  } catch { return null; }
}

/** Inflates raw DEFLATE data. */
export async function inflateRaw(u8: Uint8Array): Promise<Uint8Array> {
  const s = await viaStream(u8, 'deflate-raw');
  if (s) return s;
  const { immortalInflate } = await import('./recover/inflate');
  const r = immortalInflate(u8);
  if (r.isCorrupt && !r.data.length) throw new Error('The compressed data is damaged.');
  return r.data;
}

/** Decompresses a gzip file (all members). Damaged data yields what could be decoded. */
export async function gunzip(data: Uint8Array | ArrayBuffer): Promise<Uint8Array> {
  const u8 = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (!isGzip(u8)) throw new Error('Not a gzip file');
  const s = await viaStream(u8, 'gzip');
  if (s) return s;
  const { immortalInflate } = await import('./recover/inflate');
  const r = immortalInflate(u8.subarray(gzipDataOffset(u8, 0)));
  if (r.isCorrupt && !r.data.length) throw new Error('The compressed data is damaged.');
  return r.data;
}

/** Reads the files of a POSIX/GNU tar archive (as written inside old KOffice .tar.gz stores). */
export function untar(u8: Uint8Array): Map<string, Uint8Array> {
  const out = new Map<string, Uint8Array>();
  const dec = new TextDecoder('latin1');
  let p = 0;
  while (p + 512 <= u8.length) {
    const hdr = u8.subarray(p, p + 512);
    if (hdr.every((b) => b === 0)) break;
    const str = (o: number, n: number) => dec.decode(hdr.subarray(o, o + n)).replace(/\0.*$/s, '');
    let name = str(0, 100);
    const prefix = str(345, 155);
    if (prefix && dec.decode(hdr.subarray(257, 262)) === 'ustar') name = prefix + '/' + name;
    const size = parseInt(str(124, 12).trim() || '0', 8) || 0;
    const type = String.fromCharCode(hdr[156] || 48);
    p += 512;
    if (type === '0' || type === '\0' || type === '7') out.set(name.replace(/^\.\//, ''), u8.subarray(p, p + size));
    p += Math.ceil(size / 512) * 512;
  }
  return out;
}

export const isTar = (u8: Uint8Array): boolean => u8.length >= 512 && /^ustar/.test(String.fromCharCode(...u8.subarray(257, 262)));
