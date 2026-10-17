/**
 * First-record signatures of the Lotus-format files (Lotus 1-2-3, Symphony, Works, Quattro for
 * DOS), kept apart from the reader so the file sniffer can use them without loading it.
 */
const u16 = (b: Uint8Array, o: number): number => (o + 2 <= b.length ? b[o] | (b[o + 1] << 8) : 0);

export type LotusVariant = 'wks' | 'wk1' | 'wrk' | 'wk3' | 'wk4' | '123' | 'works-dos' | 'works-win' | 'wq1' | 'wq2';

/** The variant a Lotus-format file belongs to, from its first record; null when it is not one. */
export function lotusVariant(b: Uint8Array): LotusVariant | null {
  if (b.length < 6) return null;
  const op = u16(b, 0), len = u16(b, 2), v = u16(b, 4);
  if (op === 0x00ff && len === 2 && v === 0x0404) return 'works-win';
  if (op !== 0 || (len !== 2 && len !== 26)) return null;
  if (len === 2) {
    if (v === 0x0404) return hasWorksRecords(b) ? 'works-dos' : 'wks';
    if (v === 0x0405) return 'wrk';
    if (v === 0x0406) return 'wk1';
    if (v === 0x5120) return 'wq1';
    if (v === 0x5121) return 'wq2';
    return null;
  }
  if (v >= 0x1000 && v <= 0x1002) return 'wk3';
  if (v === 0x1003 || v === 0x1004) return 'wk4';
  if (v >= 0x1005 && v <= 0x1007) return '123';
  return null;
}

function hasWorksRecords(b: Uint8Array): boolean {
  let p = 0;
  for (let k = 0; k < 40 && p + 4 <= b.length; k++) {
    const op = u16(b, p);
    if (op >= 0x5400 && op <= 0x54ff) return true;
    p += 4 + u16(b, p + 2);
  }
  return false;
}

export const LOTUS_VARIANT_LABEL: Record<LotusVariant, string> = {
  wks: 'Lotus 1-2-3 Release 1A worksheet (WKS)', wk1: 'Lotus 1-2-3 Release 2 worksheet (WK1)', wrk: 'Symphony worksheet (WRK)',
  wk3: 'Lotus 1-2-3 Release 3 workbook (WK3)', wk4: 'Lotus 1-2-3 Release 4/5 workbook (WK4)', '123': 'Lotus 1-2-3 97/Millennium workbook (123)',
  'works-dos': 'Microsoft Works for DOS spreadsheet (WKS)', 'works-win': 'Microsoft Works for Windows spreadsheet (WKS)',
  wq1: 'Quattro / Quattro Pro for DOS spreadsheet (WQ1)', wq2: 'Quattro Pro 5 for DOS spreadsheet (WQ2)',
};

