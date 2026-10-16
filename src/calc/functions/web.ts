/**
 * Web functions. ENCODEURL and FILTERXML work offline; WEBSERVICE is not supported because cells
 * never make network requests (it returns #VALUE!).
 */
import { defFn } from './registry';
import { str, parseNumeric, ERR, CellError, type Value } from '../values';
import type { Matrix, Scalar } from '../../model/types';

const S = (v: Value): Scalar => v as Scalar;

/** Percent-encodes everything except RFC 3986 unreserved characters (A-Z a-z 0-9 - _ . ~). */
function encodeUrl(s: string): string {
  return encodeURIComponent(s).replace(/[!'()*]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase());
}

function filterXml(xml: string, xpath: string): Scalar | Matrix {
  if (typeof DOMParser === 'undefined' || typeof XPathResult === 'undefined') return ERR.value;
  let doc: Document;
  try {
    doc = new DOMParser().parseFromString(xml, 'application/xml');
  } catch {
    return ERR.value;
  }
  if (!doc.documentElement || doc.getElementsByTagName('parsererror').length) return ERR.value;
  let res: XPathResult;
  try {
    res = doc.evaluate(xpath, doc, null, XPathResult.ANY_TYPE, null);
  } catch {
    return ERR.value;
  }
  const conv = (t: string): Scalar => {
    const n = parseNumeric(t);
    return n != null && /^\s*[-+]?(\d|\.\d)/.test(t) ? n : t;
  };
  switch (res.resultType) {
    case XPathResult.NUMBER_TYPE: return res.numberValue;
    case XPathResult.STRING_TYPE: return conv(res.stringValue);
    case XPathResult.BOOLEAN_TYPE: return res.booleanValue;
    default: {
      const out: Scalar[] = [];
      for (let n = res.iterateNext(); n; n = res.iterateNext()) out.push(conv(n.textContent ?? ''));
      if (!out.length) return ERR.value;
      return out.length === 1 ? out[0] : out.map((x) => [x]);
    }
  }
}

defFn([
  {
    name: 'ENCODEURL', min: 1, max: 1, category: 'Web', desc: 'Percent-encodes text for use in a URL.', params: ['text'],
    impl: (a) => encodeUrl(str(S(a[0]))),
  },
  {
    name: 'FILTERXML', min: 2, max: 2, category: 'Web', desc: 'Extracts values from XML text using an XPath expression.', params: ['xml', 'xpath'],
    impl: (a) => filterXml(str(S(a[0])), str(S(a[1]))),
  },
  {
    name: 'WEBSERVICE', min: 1, max: 1, category: 'Web', desc: 'Returns data from a web service (not supported: cells do not access the network).', params: ['url'],
    impl: (a) => {
      str(S(a[0]));
      return new CellError('#VALUE!', 'Network functions are not supported in cells');
    },
  },
]);
