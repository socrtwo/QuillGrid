/**
 * iWork Archive (.iwa) primitives used by the Apple Numbers 2013+ reader: Snappy block
 * decompression, the iWork chunk framing, protocol buffer decoding without schemas, and the
 * object index that maps archive identifiers to their messages.
 *
 * An .iwa file is a series of chunks, each a 0x00 byte, a 24-bit little-endian length and a
 * Snappy-compressed block (no CRC). The decompressed stream is a sequence of archives: a
 * varint-prefixed ArchiveInfo message (field 1 identifier, field 2 MessageInfo entries with
 * field 1 type and field 3 length) followed by the message payloads it describes.
 */

/** Decompresses one raw Snappy block. */
export function snappyDecompress(src: Uint8Array): Uint8Array {
  let p = 0;
  let n = 0, shift = 0;
  for (;;) { const b = src[p++]; n |= (b & 0x7f) << shift; if (!(b & 0x80)) break; shift += 7; if (shift > 28) throw new Error('bad snappy length'); }
  const out = new Uint8Array(n >>> 0);
  let o = 0;
  while (p < src.length) {
    const tag = src[p++];
    const kind = tag & 3;
    if (kind === 0) {
      let len = tag >> 2;
      if (len >= 60) { const bytes = len - 59; len = 0; for (let i = 0; i < bytes; i++) len |= src[p++] << (8 * i); }
      len += 1;
      if (o + len > out.length || p + len > src.length) throw new Error('snappy literal out of range');
      out.set(src.subarray(p, p + len), o); p += len; o += len;
      continue;
    }
    let len: number, off: number;
    if (kind === 1) { len = ((tag >> 2) & 7) + 4; off = ((tag >> 5) << 8) | src[p++]; }
    else if (kind === 2) { len = (tag >> 2) + 1; off = src[p] | (src[p + 1] << 8); p += 2; }
    else { len = (tag >> 2) + 1; off = (src[p] | (src[p + 1] << 8) | (src[p + 2] << 16) | (src[p + 3] << 24)) >>> 0; p += 4; }
    if (!off || off > o || o + len > out.length) throw new Error('snappy copy out of range');
    for (let i = 0; i < len; i++, o++) out[o] = out[o - off];
  }
  return o === out.length ? out : out.subarray(0, o);
}

/** Undoes the iWork chunk framing of an .iwa file. */
export function unframeIwa(u8: Uint8Array): Uint8Array {
  const parts: Uint8Array[] = [];
  let p = 0, total = 0;
  while (p + 4 <= u8.length) {
    if (u8[p] !== 0) throw new Error('This .iwa file is not in the expected format (it may be encrypted).');
    const len = u8[p + 1] | (u8[p + 2] << 8) | (u8[p + 3] << 16);
    const block = snappyDecompress(u8.subarray(p + 4, p + 4 + len));
    parts.push(block); total += block.length;
    p += 4 + len;
  }
  const out = new Uint8Array(total);
  let o = 0;
  for (const b of parts) { out.set(b, o); o += b.length; }
  return out;
}

// --------------------------------------------------------------------------------------------
// Protocol buffers (wire format only)

export interface PbField { wire: number; /** varint value (as a number; exact below 2^53) */ v: number; /** length-delimited bytes or fixed bytes */ b?: Uint8Array }
export type PbMessage = Map<number, PbField[]>;

export function readVarint(u8: Uint8Array, p: number): [number, number] {
  let v = 0, mul = 1, i = 0;
  for (;;) {
    const b = u8[p + i];
    if (b === undefined) throw new Error('truncated varint');
    v += (b & 0x7f) * mul;
    i++;
    if (!(b & 0x80)) break;
    mul *= 128;
    if (i > 10) throw new Error('bad varint');
  }
  return [v, i];
}

/** Decodes a message into field number -> occurrences. */
export function pbDecode(u8: Uint8Array): PbMessage {
  const m: PbMessage = new Map();
  let p = 0;
  while (p < u8.length) {
    const [key, kn] = readVarint(u8, p); p += kn;
    const field = Math.floor(key / 8), wire = key & 7;
    let f: PbField;
    switch (wire) {
      case 0: { const [v, n] = readVarint(u8, p); p += n; f = { wire, v }; break; }
      case 1: f = { wire, v: 0, b: u8.subarray(p, p + 8) }; p += 8; break;
      case 2: { const [len, n] = readVarint(u8, p); p += n; f = { wire, v: len, b: u8.subarray(p, p + len) }; p += len; break; }
      case 5: f = { wire, v: 0, b: u8.subarray(p, p + 4) }; p += 4; break;
      default: throw new Error(`unsupported protobuf wire type ${wire}`);
    }
    if (p > u8.length) throw new Error('truncated protobuf message');
    let list = m.get(field);
    if (!list) { list = []; m.set(field, list); }
    list.push(f);
  }
  return m;
}

export const pbVarint = (m: PbMessage | undefined, f: number): number | undefined => m?.get(f)?.[0]?.v;
export const pbBytes = (m: PbMessage | undefined, f: number): Uint8Array | undefined => m?.get(f)?.[0]?.b;
export const pbMsg = (m: PbMessage | undefined, f: number): PbMessage | undefined => { const b = pbBytes(m, f); return b ? pbDecode(b) : undefined; };
export const pbMsgs = (m: PbMessage | undefined, f: number): PbMessage[] => (m?.get(f) || []).filter((x) => x.b).map((x) => pbDecode(x.b!));
export const pbString = (m: PbMessage | undefined, f: number): string | undefined => { const b = pbBytes(m, f); return b ? new TextDecoder().decode(b) : undefined; };
export const pbFloat = (m: PbMessage | undefined, f: number): number | undefined => { const b = pbBytes(m, f); return b && b.length === 4 ? new DataView(b.buffer, b.byteOffset, 4).getFloat32(0, true) : b && b.length === 8 ? new DataView(b.buffer, b.byteOffset, 8).getFloat64(0, true) : undefined; };
/** TSP.Reference (field 1 identifier) -> object id. */
export const pbRef = (m: PbMessage | undefined, f: number): number | undefined => pbVarint(pbMsg(m, f), 1);
export const pbRefs = (m: PbMessage | undefined, f: number): number[] => pbMsgs(m, f).map((r) => pbVarint(r, 1)).filter((x): x is number => x != null);

/** Packed or repeated varints. */
export function pbPacked(m: PbMessage | undefined, f: number): number[] {
  const out: number[] = [];
  for (const x of m?.get(f) || []) {
    if (x.wire === 0) out.push(x.v);
    else if (x.b) { let p = 0; while (p < x.b.length) { const [v, n] = readVarint(x.b, p); out.push(v); p += n; } }
  }
  return out;
}

// --------------------------------------------------------------------------------------------
// Object index

export interface IwaObject { id: number; type: number; data: Uint8Array; file: string }

/** Parses an unframed archive stream and adds its objects (first message of each archive) to the index. */
export function indexArchives(stream: Uint8Array, file: string, out: Map<number, IwaObject>): void {
  let p = 0;
  while (p < stream.length) {
    const [len, n] = readVarint(stream, p); p += n;
    const info = pbDecode(stream.subarray(p, p + len)); p += len;
    const id = pbVarint(info, 1) ?? 0;
    const msgs = pbMsgs(info, 2);
    msgs.forEach((mi, i) => {
      const type = pbVarint(mi, 1) ?? 0, length = pbVarint(mi, 3) ?? 0;
      const data = stream.subarray(p, p + length);
      p += length;
      if (i === 0 && !out.has(id)) out.set(id, { id, type, data, file });
    });
  }
}

// --------------------------------------------------------------------------------------------
// Writing (used by tests to build fixtures from the format description)

export function writeVarint(v: number, out: number[] = []): number[] {
  let x = v;
  while (x >= 0x80) { out.push((x % 128) | 0x80); x = Math.floor(x / 128); }
  out.push(x);
  return out;
}

/** Fixed-width values: { f32: n } (wire type 5) and { f64: n } (wire type 1). */
export type PbFixed = { f32: number } | { f64: number };
export type PbValue = number | string | Uint8Array | PbFixed | PbValue[] | { [field: number]: PbValue };

/** Encodes {field: value}: numbers as varints, strings and bytes as length-delimited, objects as sub-messages, arrays repeat. */
export function pbEncode(msg: { [field: number]: PbValue }): Uint8Array {
  const out: number[] = [];
  const put = (field: number, v: PbValue) => {
    if (Array.isArray(v)) { for (const x of v) put(field, x); return; }
    if (typeof v === 'number') { writeVarint(field * 8, out); writeVarint(v, out); return; }
    if (v && typeof v === 'object' && ('f32' in v || 'f64' in v)) {
      const wide = 'f64' in v;
      const dv = new DataView(new ArrayBuffer(wide ? 8 : 4));
      if (wide) dv.setFloat64(0, (v as { f64: number }).f64, true); else dv.setFloat32(0, (v as { f32: number }).f32, true);
      writeVarint(field * 8 + (wide ? 1 : 5), out); out.push(...new Uint8Array(dv.buffer));
      return;
    }
    const b = typeof v === 'string' ? new TextEncoder().encode(v) : v instanceof Uint8Array ? v : pbEncode(v as { [field: number]: PbValue });
    writeVarint(field * 8 + 2, out); writeVarint(b.length, out); out.push(...b);
  };
  for (const [k, v] of Object.entries(msg)) put(+k, v);
  return new Uint8Array(out);
}

/** Snappy "compression" with literals only (valid input for any Snappy decoder). */
export function snappyLiteral(u8: Uint8Array): Uint8Array {
  const out: number[] = writeVarint(u8.length);
  for (let p = 0; p < u8.length; p += 65536) {
    const chunk = u8.subarray(p, Math.min(u8.length, p + 65536));
    const n = chunk.length - 1;
    if (n < 60) out.push(n << 2);
    else if (n < 256) out.push(60 << 2, n);
    else out.push(61 << 2, n & 255, n >> 8);
    out.push(...chunk);
  }
  return new Uint8Array(out);
}

/** Frames archives ({id, type, message}) as an .iwa file. */
export function buildIwa(objects: { id: number; type: number; msg: Uint8Array }[]): Uint8Array {
  const stream: number[] = [];
  for (const o of objects) {
    const info = pbEncode({ 1: o.id, 2: [{ 1: o.type, 2: new Uint8Array(writeVarint(1)), 3: o.msg.length }] });
    writeVarint(info.length, stream); stream.push(...info, ...o.msg);
  }
  const block = snappyLiteral(new Uint8Array(stream));
  return new Uint8Array([0, block.length & 255, (block.length >> 8) & 255, (block.length >> 16) & 255, ...block]);
}
