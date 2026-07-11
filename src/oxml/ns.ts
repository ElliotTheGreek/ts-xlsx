/**
 * Namespace registry — port of openpyxl/xml/constants.py.
 *
 * Prefixes here are the *canonical* ones used when ts-xlsx creates elements.
 * SpreadsheetML's main namespace is normally the *default* namespace in a
 * real workbook (`<worksheet xmlns="…/spreadsheetml/2006/main">`), so elements
 * created inside an existing tree come out unprefixed (createElement resolves
 * the in-scope default declaration). Matching is always by namespace URI +
 * local name, never by prefix string.
 */
export const nsmap = {
  /** SpreadsheetML main (worksheet/workbook/sharedStrings/styles). */
  main: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  /** officeDocument relationships (r:id references inside parts). */
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  /** Package relationships (.rels file vocabulary). */
  pr: "http://schemas.openxmlformats.org/package/2006/relationships",
  /** [Content_Types].xml vocabulary. */
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
  /** Core document properties. */
  cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  dc: "http://purl.org/dc/elements/1.1/",
  dcterms: "http://purl.org/dc/terms/",
  dcmitype: "http://purl.org/dc/dcmitype/",
  /** Extended (app) document properties + variant types. */
  ep: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
  vt: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
  /** DrawingML main + chart + spreadsheet drawing. */
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  /** Markup compatibility (mc:Ignorable / AlternateContent). */
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  /** Microsoft spreadsheet extensions (x14ac row/col, x14 rules, xm refs). */
  x14: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
  x14ac: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
  xm: "http://schemas.microsoft.com/office/excel/2006/main",
  xr: "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
  xml: "http://www.w3.org/XML/1998/namespace",
  xsi: "http://www.w3.org/2001/XMLSchema-instance",
} as const;

export type NsPrefix = keyof typeof nsmap;

/** A canonical namespace-prefixed tag like "main:c" or "r:id". Unprefixed
 * names in .rels / [Content_Types].xml are addressed via their default-ns URI. */
export type NsTag = `${NsPrefix}:${string}`;

export function nsUriOf(prefix: NsPrefix): string {
  return nsmap[prefix];
}

/** Split "main:c" into { nsUri, localName, prefix }. */
export function splitNsTag(tag: NsTag): { nsUri: string; localName: string; prefix: NsPrefix } {
  const c = tag.indexOf(":");
  const prefix = tag.slice(0, c) as NsPrefix;
  const uri = nsmap[prefix];
  if (uri === undefined) throw new Error(`unknown canonical namespace prefix: ${tag}`);
  return { nsUri: uri, localName: tag.slice(c + 1), prefix };
}

/** `nsdecls("main", "r")` → `xmlns="…" xmlns:r="…"` (the `main` prefix emits
 * a default-namespace declaration, matching how Excel authors the parts). */
export function nsdecls(...prefixes: NsPrefix[]): string {
  return prefixes
    .map((p) => (p === "main" ? `xmlns="${nsmap[p]}"` : `xmlns:${p}="${nsmap[p]}"`))
    .join(" ");
}
