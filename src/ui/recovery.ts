/** Recovery report dialog and the manual recovery guide (adapted from socrtwo/wordrecovery-SF). */
import { h, icon, pickFile } from './dom';
import { openDialog, toast } from './dialog';
import type { RecoveryReport, MethodStatus } from '../io/recover';
import type { Workbook } from '../model/workbook';
import { MIME } from '../io/recover/fileid';
import { saveAs, openExternal, platformName } from '../platform/native';

const STATUS: Record<MethodStatus, [string, string]> = {
  success: ['check_circle', 'Succeeded'],
  partial: ['adjust', 'Partial'],
  failed: ['cancel', 'Failed'],
  skipped: ['remove_circle_outline', 'Skipped'],
};

export interface ReportActions {
  /** Open the recovered workbook in a new tab. */
  openBook?: (book: Workbook, name: string, report: RecoveryReport) => void;
  /** The recovered workbook is already open: switch to it. */
  open?: () => void;
}

export function recoverySummary(r: RecoveryReport): string {
  return r.summary || (r.book ? 'The workbook was recovered.' : 'Nothing could be recovered from this file.');
}

const plural = (n: number, w: string) => `${n.toLocaleString('en-US')} ${n === 1 ? w : w + 's'}`;

function sheetRows(r: RecoveryReport): HTMLElement | null {
  if (!r.book) return null;
  return h('table', { class: 'ql-recovery-methods' },
    h('tbody', null, r.book.sheets.map((sh) => {
      let cells = 0, formulas = 0;
      for (const c of sh.cells.values()) { if ((c.v != null && c.v !== '') || c.f) cells++; if (c.f) formulas++; }
      const unverified = /\(unverified/.test(sh.name);
      return h('tr', { class: unverified ? 'ql-st-partial' : 'ql-st-success' },
        h('td', { title: unverified ? 'Unverified' : 'Recovered' }, icon(unverified ? 'warning' : 'grid_on')),
        h('td', null, h('b', null, sh.name), h('div', { class: 'ql-help' }, `${plural(cells, 'cell')}${formulas ? `, ${plural(formulas, 'formula')}` : ''}${sh.visibility !== 'visible' ? ' (hidden)' : ''}`)));
    })));
}

export async function showRecoveryReport(r: RecoveryReport, actions: ReportActions = {}): Promise<void> {
  const base = r.fileName.replace(/\.[^.]+$/, '');
  const body = h('div', { class: 'ql-recovery' },
    h('p', { class: 'ql-recovery-summary' }, h('b', null, recoverySummary(r))),
    h('p', { class: 'ql-help' }, `Identified as: ${r.identification.description} (${r.identification.confidence} confidence).`),
    r.unverified ? h('p', { class: 'ql-help', style: 'color:#c62828' }, icon('warning'), ' Some of the recovered data was decoded from a damaged area of the file. Check it against another copy before relying on it.') : null,
    r.textOnly ? h('p', { class: 'ql-help', style: 'color:#b26a00' }, 'Only loose text was recovered: each line is one row, and formatting, formulas and cell positions are lost.') : null,
    r.notes.length ? h('ul', { class: 'ql-recovery-notes' }, r.notes.map((n) => h('li', null, n))) : null,
    h('h3', null, 'Recovery methods'),
    h('table', { class: 'ql-recovery-methods' },
      h('tbody', null, r.methods.map((m) => {
        const [ic, label] = STATUS[m.status];
        const best = r.bestMethod === m.id;
        return h('tr', { class: `ql-st-${m.status}${best ? ' ql-best' : ''}` },
          h('td', { title: label }, icon(ic)),
          h('td', null, h('b', null, m.name), best ? h('span', { class: 'ql-chip' }, 'used') : null,
            h('div', { class: 'ql-help' }, m.detail || label), h('div', { class: 'ql-help ql-dim' }, m.desc)));
      }))),
    r.book ? h('div', null, h('h3', null, 'Recovered sheets'), sheetRows(r)) : null,
    r.repaired && (r.repaired.repairedParts.length || r.repaired.salvagedParts.length || r.repaired.lostParts.length)
      ? h('div', null,
        h('h3', null, 'Package parts'),
        r.repaired.salvagedParts.length ? h('p', { class: 'ql-help' }, `Salvaged from raw bytes: ${r.repaired.salvagedParts.join(', ')}`) : null,
        r.repaired.repairedParts.length ? h('p', { class: 'ql-help' }, `Repaired: ${r.repaired.repairedParts.join('; ')}`) : null,
        r.repaired.lostParts.length ? h('p', { class: 'ql-help' }, `Could not recover: ${r.repaired.lostParts.join(', ')}`) : null)
      : null,
    r.foreign.length ? h('div', null,
      h('h3', null, 'Other data found inside the file'),
      r.foreign.map((seg) => h('div', { class: 'ql-recovery-seg' },
        h('span', null, `${seg.description}: ${fmtSize(seg.length)} at byte ${seg.offset.toLocaleString('en-US')}`),
        h('button', { class: 'ql-btn', onclick: () => saveAs(new Blob([seg.bytes.slice()], { type: MIME[seg.ext] || 'application/octet-stream' }), seg.suggestedName || `${base}.part.${seg.ext}`) }, 'Save piece…'),
        seg.recommended?.[0] ? h('button', { class: 'ql-link', onclick: () => openExternal(seg.recommended![0].url) }, `Repair with ${seg.recommended[0].name}`) : null)))
      : null,
    !r.book && r.recommended.length ? h('div', null,
      h('h3', null, 'Other tools'),
      h('p', { class: 'ql-help' }, 'These free S2 recovery tools specialize in this kind of file:'),
      h('ul', { class: 'ql-recovery-notes' }, r.recommended.slice(0, 4).map((p) => h('li', null, h('button', { class: 'ql-link', onclick: () => openExternal(p.url) }, p.name)))))
      : null,
  );
  const buttons: { label: string; onClick: (v: Record<string, any>, ctl: any) => void; primary?: boolean }[] = [];
  if (r.book && (actions.openBook || actions.open)) {
    buttons.push({
      label: 'Open recovered workbook', primary: true,
      onClick: (_v, ctl) => { ctl.close(null); if (actions.openBook) actions.openBook(r.book!, `${base} (Recovered)`, r); else actions.open!(); },
    });
  }
  if (r.book && !actions.openBook && !actions.open) {
    buttons.push({
      label: 'Save recovered workbook…',
      onClick: async () => {
        try {
          const { exportFile } = await import('../io');
          saveAs(await exportFile(r.book!, 'xlsx'), `${base} (recovered).xlsx`);
        } catch (e: any) { toast(`Could not save: ${e?.message || e}`); }
      },
    });
  }
  if (r.repaired) {
    const kind = r.repaired.kind;
    buttons.push({
      label: `Save repaired package (.${kind})…`,
      onClick: () => saveAs(r.repaired!.blob, r.repaired!.name),
    });
  }
  buttons.push({
    label: 'Copy report',
    onClick: async () => {
      try {
        const { reportText } = await import('../io/recover');
        await navigator.clipboard.writeText(reportText(r));
        toast('Recovery report copied');
      } catch { toast('Clipboard is not available'); }
    },
  });
  buttons.push({ label: 'Manual recovery options', onClick: () => { void showManualRecovery(actions); } });
  await openDialog({ title: `Recovery Report: ${r.fileName}`, content: body, wide: true, ok: 'Close', cancel: false, buttons });
}

function fmtSize(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1048576) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1048576).toFixed(1)} MB`;
}

const link = (url: string, text: string) => `<a href="${url}" data-external="1">${text}</a>`;

/** Per-platform manual recovery guide for workbooks, adapted from the Word Recovery "Manual recovery" tab. */
const MANUAL: Record<string, [string, string]> = {
  windows: ['Windows', `
    <h3>1. AutoRecover and unsaved workbooks</h3>
    <p>Excel saves AutoRecover copies (<code>.xar</code> files, which are ordinary workbooks; rename them to <code>.xlsx</code> or <code>.xls</code> to open them here). Paste into the Explorer address bar:</p>
    <ul><li><code>%AppData%\\Microsoft\\Excel\\</code></li><li><code>%LocalAppData%\\Microsoft\\Office\\UnsavedFiles\\</code></li></ul>
    <p>In Excel: File &gt; Info &gt; Manage Workbook &gt; <b>Recover Unsaved Workbooks</b>.</p>
    <h3>2. Temporary and backup files</h3>
    <p><code>%Temp%</code> and the workbook's folder: look for files with no extension and 8-character hex names (Excel's save temp files), <code>~$name.xlsx</code> owner files (not data), and <code>Backup of name.xlk</code> (created when "Always create backup" is on). Open them with <b>Open and Repair</b>.</p>
    <h3>3. Previous Versions / Shadow Copies</h3>
    <p>Right-click the file or folder &gt; Properties &gt; <b>Previous Versions</b> (needs System Protection or File History). OneDrive and SharePoint keep a version history: right-click the file &gt; <b>Version history</b>.</p>
    <h3>4. Excel's own repair</h3>
    <p>File &gt; Open &gt; Browse, select the file, click the arrow on <b>Open</b> &gt; <b>Open and Repair</b> &gt; <b>Repair</b>, then <b>Extract Data</b> (values or formulas) if repair fails. Setting calculation to manual first (File &gt; Options &gt; Formulas) helps with workbooks that hang on open.</p>
    <h3>5. Microsoft's guidance</h3>
    <p>${link('https://support.microsoft.com/en-us/office/repair-a-corrupted-workbook-153a45f4-6cab-44b1-93ca-801ddcd4ea53', 'Repair a corrupted workbook')} · ${link('https://support.microsoft.com/en-us/office/recover-an-earlier-version-of-an-office-file-851b6ea8-1d12-43e2-b3f2-1095bb16cee2', 'Recover an earlier version of an Office file')}</p>
    <h3>6. Deleted-file recovery</h3>
    <p>${link('https://www.ccleaner.com/recuva', 'Recuva')} · ${link('https://www.cgsecurity.org/wiki/PhotoRec', 'PhotoRec')} · ${link('https://www.voidtools.com/', 'Everything')}</p>`],
  macos: ['macOS', `
    <h3>1. AutoRecovery files</h3>
    <ul><li><code>~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/</code></li><li><code>~/Library/Containers/com.microsoft.Excel/Data/Library/Preferences/AutoRecovery/</code></li></ul>
    <p>Finder &gt; Go &gt; <b>Go to Folder…</b>; hold Option on the Go menu to show Library.</p>
    <h3>2. Temporary files</h3><p><code>/private/var/folders/</code> (Excel's scratch directory) and <code>/tmp/</code>.</p>
    <h3>3. Time Machine and iCloud</h3><p>Open the containing folder and enter Time Machine, or use File &gt; Revert To &gt; Browse All Versions in Numbers for iCloud files.</p>
    <h3>4. Open and Repair</h3><p>Excel 2016 and later: File &gt; Open, select the file, arrow on Open &gt; <b>Open and Repair</b>. ${link('https://www.libreoffice.org/', 'LibreOffice Calc')} and Numbers are often more tolerant.</p>`],
  linux: ['Linux', `
    <h3>1. LibreOffice Calc</h3><p>${link('https://www.libreoffice.org/', 'LibreOffice Calc')} tolerates many damaged XLSX/XLS/ODS files. Headless conversion: <code>soffice --headless --convert-to xlsx damaged.xls</code>.</p>
    <h3>2. Repair the ZIP container</h3><pre><code>cp damaged.xlsx broken.zip
zip -FF broken.zip --out fixed.zip
mv fixed.zip recovered.xlsx</code></pre>
    <h3>3. Extract values directly</h3><pre><code>unzip -p damaged.xlsx xl/sharedStrings.xml | sed 's/&lt;[^&gt;]*&gt;/ /g'
unzip -p damaged.xlsx xl/worksheets/sheet1.xml | grep -o '&lt;v&gt;[^&lt;]*'
ssconvert damaged.xls recovered.csv   # Gnumeric</code></pre>
    <h3>4. Strings on .xls binaries</h3><pre><code>strings -n 4 damaged.xls | less
strings -n 4 -e l damaged.xls   # UTF-16 shared strings</code></pre>
    <h3>5. Deleted-file recovery</h3><p>${link('https://www.cgsecurity.org/wiki/PhotoRec', 'PhotoRec')} · ${link('https://github.com/sleuthkit/sleuthkit', 'The Sleuth Kit')}</p>`],
  android: ['Android', `
    <h3>1. Version history in the app that created the file</h3>
    <p>Excel for Android: open the workbook, … menu &gt; <b>History</b>. Google Sheets/Drive and OneDrive keep versions too.</p>
    <h3>2. Trash and recently deleted</h3>
    <p>Google Drive &gt; <b>Trash</b>, OneDrive &gt; <b>Recycle bin</b>, and the Files app's trash keep deleted files for about 30 days.</p>
    <h3>3. QuillGrid's own copies</h3>
    <p>Open &gt; Recent keeps a cached copy of each file you opened here.</p>
    <h3>4. Online viewers</h3>
    <p>${link('https://www.office.com/launch/excel', 'Excel for the web')} and ${link('https://sheets.google.com/', 'Google Sheets')} sometimes open what desktop apps refuse.</p>`],
  ios: ['iOS / iPadOS', `
    <h3>1. Version history</h3><p>Excel for iOS: … menu &gt; <b>Restore</b> &gt; <b>Browse Version History</b>.</p>
    <h3>2. Files &gt; Recently Deleted</h3><p>Files keeps deleted items for about 30 days.</p>
    <h3>3. Online viewers</h3><p>${link('https://www.office.com/launch/excel', 'Excel for the web')} · ${link('https://sheets.google.com/', 'Google Sheets')}</p>`],
  tools: ['More S2 recovery tools', `
    <p>Free tools by the author of QuillGrid's recovery engine, each specialized for one kind of damage:</p>
    <ul>
      <li>${link('https://socrtwo.github.io/corruptexcelrec-SF/', 'S2 Recovery Tools for Microsoft Excel')} · ${link('https://socrtwo.github.io/excelrcvryaddin-SF/', 'Excel Recovery Addin')}</li>
      <li>${link('https://socrtwo.github.io/Immortal-Zip/', 'Immortal Zip')} · ${link('https://socrtwo.github.io/xmltrncatorfixr-SF/app/', 'XML Truncator-Fixer')} · ${link('https://socrtwo.github.io/Universal-File-Repair-Tool/', 'Universal File Repair Tool')}</li>
      <li>${link('https://socrtwo.github.io/oorecovery-SF/', 'OO Recovery (OpenDocument)')} · ${link('https://socrtwo.github.io/savvyoffice-SF/', 'Savvy Repair for Microsoft Office')} · ${link('https://socrtwo.github.io/wordrecovery-SF/', 'Word Recovery')}</li>
    </ul>`],
};

/** Pick a file, run every recovery method on it (Open and Repair), and show the report. */
export async function recoverPickedFile(actions: ReportActions = {}): Promise<RecoveryReport | null> {
  const [file] = await pickFile('*/*');
  if (!file) return null;
  const { recoverWorkbook } = await import('../io/recover');
  toast(`Recovering ${file.name}…`, { ms: 2500 });
  const report = await recoverWorkbook(await file.arrayBuffer(), file.name, { forceRepair: true });
  await showRecoveryReport(report, actions);
  return report;
}

export async function showManualRecovery(actions: ReportActions = {}): Promise<void> {
  const plat = platformName();
  const guess = plat === 'android' ? 'android' : plat === 'ios' ? 'ios' : /Mac/.test(navigator.platform) ? 'macos' : /Linux/.test(navigator.platform) ? 'linux' : 'windows';
  const content = h('div', { class: 'ql-manual' });
  const pane = h('div', { class: 'ql-manual-body' });
  const tabs = h('div', { class: 'ql-seg' }, Object.entries(MANUAL).map(([k, [label]]) => h('button', { class: 'ql-seg-btn', 'data-k': k, onclick: () => show(k) }, label)));
  const show = (k: string) => {
    pane.innerHTML = MANUAL[k][1];
    tabs.querySelectorAll('button').forEach((b) => b.classList.toggle('ql-on', b.getAttribute('data-k') === k));
  };
  pane.addEventListener('click', (e) => {
    const a = (e.target as HTMLElement).closest('a[data-external]') as HTMLAnchorElement | null;
    if (a) { e.preventDefault(); openExternal(a.href); }
  });
  const intro = h('p', { class: 'ql-help' }, 'QuillGrid can try every recovery method on a damaged workbook (XLSX, XLSM, XLS, ODS, CSV, HTML or XML) or on any file that may hold spreadsheet data: pick it with "Recover a file…". The steps below find other copies of the file.');
  content.append(intro, tabs, pane);
  show(guess);
  await openDialog({
    title: 'Manual Recovery Options', content, wide: true, ok: 'Close', cancel: false,
    buttons: [{ label: 'Recover a file…', primary: true, onClick: (_v, ctl) => { ctl.close(null); void recoverPickedFile(actions).catch((e) => toast(`Recovery failed: ${e?.message || e}`)); } }],
  });
}
