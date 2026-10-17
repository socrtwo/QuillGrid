/**
 * StarOffice 6 / OpenOffice.org 1.x Calc (.sxc, templates .stc): the XML format that became
 * OpenDocument. The package has the same parts (content.xml, styles.xml, meta.xml,
 * settings.xml, Pictures/) but uses the http://openoffice.org/2000/* namespaces, keeps all
 * formatting properties in one <style:properties> element, stores cell values in table:*
 * attributes, names fonts with <style:font-decl>, and puts the tables straight into
 * <office:body>.
 *
 * The documents are converted to their OpenDocument equivalents in the DOM and then read by
 * the OpenDocument importer, so styles, number formats, merges, comments, links, names, column
 * widths, row heights, frozen panes and pictures read the same way. Formulas use the bracketed
 * reference syntax ("=SUM([.B2:.B4])") that OpenDocument kept.
 */
import JSZip from 'jszip';
import type { Workbook } from '../model/workbook';
import { parseXml, decodeText, kids } from './index';
import { readOdf } from './ods';

const ODF = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  number: 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  config: 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
  chart: 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
  dc: 'http://purl.org/dc/elements/1.1/',
  xlink: 'http://www.w3.org/1999/xlink',
};

/** OpenOffice.org 1.x namespace URI -> OpenDocument URI. */
const NS_MAP: [string, string][] = [
  ['http://openoffice.org/2000/office', ODF.office],
  ['http://openoffice.org/2000/style', ODF.style],
  ['http://openoffice.org/2000/text', ODF.text],
  ['http://openoffice.org/2000/table', ODF.table],
  ['http://openoffice.org/2000/drawing', ODF.draw],
  ['http://www.w3.org/1999/XSL/Format', ODF.fo],
  ['http://openoffice.org/2000/meta', ODF.meta],
  ['http://openoffice.org/2000/datastyle', ODF.number],
  ['http://www.w3.org/2000/svg', ODF.svg],
  ['http://openoffice.org/2001/config', ODF.config],
  ['http://openoffice.org/2000/chart', ODF.chart],
];

/** Rewrites the namespace declarations of an OpenOffice.org 1.x part to OpenDocument ones. */
export function mapNamespaces(xml: string): string {
  let s = xml.replace(/<!DOCTYPE[^>[]*(\[[^\]]*\])?\s*>/, '');
  for (const [from, to] of NS_MAP) s = s.split(`"${from}"`).join(`"${to}"`);
  return s;
}

const VALUE_ATTRS = ['value-type', 'value', 'date-value', 'time-value', 'boolean-value', 'string-value', 'currency'];

/** Which *-properties elements a style family's <style:properties> stands for. */
function propertyKinds(family: string | null, parent: Element): string[] {
  switch (family) {
    case 'table-cell': return ['table-cell-properties', 'text-properties', 'paragraph-properties'];
    case 'table-column': return ['table-column-properties'];
    case 'table-row': return ['table-row-properties'];
    case 'table': return ['table-properties'];
    case 'paragraph': return ['paragraph-properties', 'text-properties'];
    case 'text': return ['text-properties'];
    case 'graphics': case 'graphic': return ['graphic-properties', 'text-properties', 'paragraph-properties'];
  }
  // Data styles (number:*-style) hold text properties.
  return parent.namespaceURI === ODF.number ? ['text-properties'] : ['text-properties', 'paragraph-properties'];
}

/** Converts an OpenOffice.org 1.x DOM (already in OpenDocument namespaces) in place. */
export function upgradeDocument(doc: Document): void {
  const all = Array.from(doc.getElementsByTagName('*'));
  for (const el of all) {
    const ns = el.namespaceURI, ln = el.localName;
    if (ns === ODF.table && (ln === 'table-cell' || ln === 'covered-table-cell')) {
      for (const a of VALUE_ATTRS) {
        const v = el.getAttributeNS(ODF.table, a);
        if (v != null && !el.hasAttributeNS(ODF.office, a)) { el.setAttributeNS(ODF.office, `office:${a}`, v); el.removeAttributeNS(ODF.table, a); }
      }
    } else if (ns === ODF.style && ln === 'properties') {
      const parent = el.parentElement;
      if (!parent) continue;
      const family = parent.getAttributeNS(ODF.style, 'family');
      for (const kind of propertyKinds(family, parent)) {
        const p = doc.createElementNS(ODF.style, `style:${kind}`);
        for (const at of Array.from(el.attributes)) if (at.namespaceURI !== 'http://www.w3.org/2000/xmlns/') p.setAttributeNS(at.namespaceURI, at.name, at.value);
        for (const ch of Array.from(el.childNodes)) p.appendChild(ch.cloneNode(true));
        parent.insertBefore(p, el);
      }
      parent.removeChild(el);
    } else if (ns === ODF.style && ln === 'font-decl') {
      const ff = doc.createElementNS(ODF.style, 'style:font-face');
      for (const at of Array.from(el.attributes)) if (at.namespaceURI !== 'http://www.w3.org/2000/xmlns/') ff.setAttributeNS(at.namespaceURI, at.name, at.value);
      const fam = el.getAttributeNS(ODF.fo, 'font-family');
      if (fam) ff.setAttributeNS(ODF.svg, 'svg:font-family', fam);
      el.parentNode!.replaceChild(ff, el);
    } else if (ns === ODF.office && ln === 'body' && !kids(el, 'spreadsheet').length) {
      const sp = doc.createElementNS(ODF.office, 'office:spreadsheet');
      while (el.firstChild) sp.appendChild(el.firstChild);
      el.appendChild(sp);
    } else if (ns === ODF.office && ln === 'annotation') {
      const ren = (from: string, to: string) => {
        const a = el.getAttributeNS(ODF.office, from);
        if (a == null) return;
        const e = doc.createElementNS(ODF.dc, `dc:${to}`);
        e.textContent = a;
        el.insertBefore(e, el.firstChild);
      };
      ren('create-date', 'date');
      ren('author', 'creator');
    } else if (ns === ODF.config && ln === 'config-item-set') {
      const n = el.getAttributeNS(ODF.config, 'name');
      if (n === 'view-settings' || n === 'configuration-settings') el.setAttributeNS(ODF.config, 'config:name', `ooo:${n}`);
    } else if (ns === ODF.draw && ln === 'image') {
      const href = el.getAttributeNS(ODF.xlink, 'href');
      if (href && href.startsWith('#')) el.setAttributeNS(ODF.xlink, 'xlink:href', href.slice(1));
    }
  }
}

function readPart(text: string | null, what: string): Document | null {
  if (text == null) return null;
  const doc = parseXml(mapNamespaces(text), what);
  upgradeDocument(doc);
  return doc;
}

export async function importSxc(data: ArrayBuffer, warnings: string[] = []): Promise<Workbook> {
  const zip = await JSZip.loadAsync(data);
  const text = async (p: string) => { const f = zip.file(p); return f ? decodeText(await f.async('uint8array')) : null; };
  const content = readPart(await text('content.xml'), 'content.xml');
  if (!content) throw new Error('content.xml is missing from this OpenOffice.org file');
  const soft = async (p: string) => { try { return readPart(await text(p), p); } catch { warnings.push(`${p} is damaged and was skipped.`); return null; } };
  const book = await readOdf({ content, styles: await soft('styles.xml'), settings: await soft('settings.xml'), meta: await soft('meta.xml'), files: zip }, warnings);
  return book;
}
