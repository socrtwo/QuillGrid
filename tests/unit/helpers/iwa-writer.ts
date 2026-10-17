/**
 * Test helper: writes a small Apple Numbers 2013+ document (iWork archives with version 5 cell
 * storage) following the published descriptions of the format, to exercise the reader without
 * Apple's software.
 */
import JSZip from 'jszip';
import { pbEncode, buildIwa, type PbValue } from '../../../src/io/iwa';

type CellSpec = { t: 'n'; v: number } | { t: 'd128'; v: number } | { t: 's'; v: string } | { t: 'date'; v: number } | { t: 'b'; v: boolean } | { t: 'f'; v: number } | { t: 'du'; v: number };

function decimal128(v: number): number[] {
  // mantissa * 10^(exponent - 6176); up to 6 decimals is enough for test values.
  let scale = 0;
  while (scale < 6 && Math.round(v * 10 ** scale) !== v * 10 ** scale) scale++;
  let mant = Math.round(Math.abs(v) * 10 ** scale);
  const exp = 6176 - scale;
  const b = new Array(16).fill(0);
  for (let i = 0; i < 14; i++) { b[i] = mant % 256; mant = Math.floor(mant / 256); }
  b[14] = ((exp & 0x7f) << 1) | (mant & 1);
  b[15] = ((exp >> 7) & 0x7f) | (v < 0 ? 0x80 : 0);
  return b;
}

function cellBytes(c: CellSpec, strings: string[]): number[] {
  const f64 = (n: number) => { const d = new DataView(new ArrayBuffer(8)); d.setFloat64(0, n, true); return [...new Uint8Array(d.buffer)]; };
  const u32 = (n: number) => [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255];
  const head = (type: number, flags: number) => [5, type, 0, 0, 0, 0, 0, 0, ...u32(flags)];
  switch (c.t) {
    case 'n': return [...head(2, 0x2), ...f64(c.v)];
    case 'd128': return [...head(2, 0x1), ...decimal128(c.v)];
    case 's': { let i = strings.indexOf(c.v); if (i < 0) { strings.push(c.v); i = strings.length - 1; } return [...head(3, 0x8), ...u32(i + 1)]; }
    case 'date': return [...head(5, 0x4), ...f64(c.v)];
    case 'b': return [...head(6, 0x2), ...f64(c.v ? 1 : 0)];
    case 'f': return [...head(2, 0x2 | 0x20 | 0x200), ...f64(c.v), ...u32(7), ...u32(1)];
    case 'du': return [...head(7, 0x2), ...f64(c.v)];
  }
}

export async function buildNumbers(): Promise<ArrayBuffer> {
  const rows: CellSpec[][] = [
    [{ t: 's', v: 'Item' }, { t: 's', v: 'Qty' }, { t: 's', v: 'Price' }, { t: 's', v: 'Total' }],
    [{ t: 's', v: 'Apples' }, { t: 'n', v: 3 }, { t: 'd128', v: 1.25 }, { t: 'f', v: 3.75 }],
    [{ t: 's', v: 'Pears' }, { t: 'n', v: 5 }, { t: 'd128', v: -2.5 }, { t: 'f', v: -12.5 }],
    [{ t: 's', v: 'When' }, { t: 'date', v: 732153600 }, { t: 'b', v: true }, { t: 'du', v: 5400 }],
  ];
  const strings: string[] = [];
  const rowInfos = rows.map((row, r) => {
    const storage: number[] = [];
    const offs: number[] = new Array(255).fill(0xffff);
    row.forEach((c, i) => { offs[i] = storage.length / 4; storage.push(...cellBytes(c, strings)); while (storage.length % 4) storage.push(0); });
    const offBytes = new Uint8Array(offs.length * 2);
    offs.forEach((o, i) => { offBytes[i * 2] = o & 255; offBytes[i * 2 + 1] = o >> 8; });
    return { 1: r, 2: row.length, 5: 5, 6: new Uint8Array(storage), 7: offBytes, 8: 1 } as PbValue;
  });
  const ref = (id: number) => ({ 1: id });
  const objects = [
    { id: 1, type: 1, msg: pbEncode({ 1: [ref(10)] }) },
    { id: 10, type: 2, msg: pbEncode({ 1: 'Sheet 1', 2: [ref(20)] }) },
    { id: 20, type: 6000, msg: pbEncode({ 2: ref(30) }) },
    { id: 30, type: 6001, msg: pbEncode({ 1: 'T-1', 4: { 1: { 2: [ref(60)] }, 2: ref(61), 3: { 1: [{ 1: 0, 2: ref(40) }] }, 4: ref(50) }, 6: rows.length, 7: 4, 8: 'Fruit stand' }) },
  ];
  const tiles = [{ id: 40, type: 6002, msg: pbEncode({ 1: 3, 2: rows.length - 1, 3: 16, 4: rows.length, 5: rowInfos }) }];
  const lists = [
    { id: 50, type: 6005, msg: pbEncode({ 1: 1, 2: strings.length + 1, 3: strings.map((s, i) => ({ 1: i + 1, 2: 1, 3: s })) }) },
    { id: 60, type: 6006, msg: pbEncode({ 1: 1, 2: [{ 1: 0, 2: { f32: 27 }, 3: 0, 4: 4 }] }) },
    { id: 61, type: 6006, msg: pbEncode({ 1: 1, 2: [{ 1: 0, 2: { f32: 120 }, 3: 0, 4: 4 }, { 1: 3, 2: { f32: 60 }, 3: 1, 4: 4 }] }) },
  ];
  const zip = new JSZip();
  zip.file('Index/Document.iwa', buildIwa(objects));
  zip.file('Index/Tables/Tile.iwa', buildIwa(tiles));
  zip.file('Index/Tables/DataList.iwa', buildIwa(lists));
  zip.file('Metadata/Properties.plist', 'bplist00');
  return zip.generateAsync({ type: 'arraybuffer' });
}
