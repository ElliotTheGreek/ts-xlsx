/**
 * Stylesheet element wrappers — port of openpyxl/styles/stylesheet.py.
 *
 * M1 wires the parts needed to resolve a cell's *number format* (custom
 * `<numFmts>` + the `<cellXfs>` chain a cell's `@s` indexes into). M3 extends
 * this same wrapper with fonts/fills/borders/alignment and the find-or-extend
 * write path — nothing here is throwaway.
 */
import { XmlElement } from "../xml/dom.js";
import { OxmlWrapper, createElement } from "./base.js";
import { nsmap, NsTag } from "./ns.js";
import { InvalidXmlError } from "../exc.js";
import { BUILTIN_FORMATS, FORMAT_GENERAL, builtinFormatId } from "../numberFormats.js";
import {
  Alignment,
  Border,
  Fill,
  Font,
  PatternFill,
  Protection,
  fillFromElement,
} from "../styles/values.js";

const MAIN = nsmap.main;

/** Overrides applied to a base xf when finding-or-adding a cell format. */
export interface XfOverrides {
  numFmtId?: number;
  fontId?: number;
  fillId?: number;
  borderId?: number;
  alignment?: Alignment;
  protection?: Protection;
}

/** Child order under `<styleSheet>` (ECMA-376 §18.8.39), for insertion. */
const STYLESHEET_CHILDREN = [
  "numFmts",
  "fonts",
  "fills",
  "borders",
  "cellStyleXfs",
  "cellXfs",
  "cellStyles",
  "dxfs",
  "tableStyles",
  "colors",
  "extLst",
] as const;

/** `<xf>` cell format record inside `<cellXfs>`. */
export class CT_Xf extends OxmlWrapper {
  get numFmtId(): number {
    const v = this.el.getAttr("numFmtId");
    return v === null ? 0 : Number(v);
  }
}

/** `<styleSheet>` root. */
export class CT_Stylesheet extends OxmlWrapper {
  /** Custom number formats: id → format code (ids >= 164). */
  get customNumFmts(): Map<number, string> {
    const out = new Map<number, string>();
    const numFmts = this.el.find(MAIN, "numFmts");
    if (numFmts === null) return out;
    for (const nf of numFmts.findAll(MAIN, "numFmt")) {
      const id = nf.getAttr("numFmtId");
      const code = nf.getAttr("formatCode");
      if (id !== null && code !== null) out.set(Number(id), code);
    }
    return out;
  }

  private get cellXfsEl(): XmlElement | null {
    return this.el.find(MAIN, "cellXfs");
  }

  get cellXfsLst(): CT_Xf[] {
    const el = this.cellXfsEl;
    if (el === null) return [];
    return el.findAll(MAIN, "xf").map((e) => new CT_Xf(e));
  }

  /** Resolve the format code applied by cell style index `s`. */
  formatCodeForStyle(s: number): string {
    const xf = this.cellXfsLst[s];
    if (xf === undefined) return FORMAT_GENERAL;
    return this.formatCodeForId(xf.numFmtId);
  }

  /** Resolve a numFmtId to its format code (custom overrides built-ins). */
  formatCodeForId(id: number): string {
    const custom = this.customNumFmts.get(id);
    if (custom !== undefined) return custom;
    return BUILTIN_FORMATS[id] ?? FORMAT_GENERAL;
  }

  // -- write path: the shared cellXfs-reuse engine ----------------------

  /** Insert a direct child in stylesheet schema order. */
  private insertOrdered(child: XmlElement, localName: string): void {
    const i = STYLESHEET_CHILDREN.indexOf(localName as (typeof STYLESHEET_CHILDREN)[number]);
    const successors = STYLESHEET_CHILDREN.slice(i + 1);
    for (const succ of successors) {
      const ref = this.el.find(MAIN, succ);
      if (ref !== null) {
        this.el.insertBefore(child, ref);
        return;
      }
    }
    this.el.appendChild(child);
  }

  /** Resolve a format code to a numFmtId, adding a custom `<numFmt>` if needed. */
  getOrAddNumFmt(code: string): number {
    const builtin = builtinFormatId(code);
    if (builtin !== undefined) return builtin;
    for (const [id, c] of this.customNumFmts) if (c === code) return id;

    let numFmts = this.el.find(MAIN, "numFmts");
    if (numFmts === null) {
      numFmts = createElement("main:numFmts", this.el);
      this.insertOrdered(numFmts, "numFmts");
    }
    let maxId = 163; // custom ids start at 164
    for (const id of this.customNumFmts.keys()) maxId = Math.max(maxId, id);
    const newId = maxId + 1;
    const nf = createElement("main:numFmt", numFmts);
    nf.setAttr("numFmtId", String(newId));
    nf.setAttr("formatCode", code);
    numFmts.appendChild(nf);
    numFmts.setAttr("count", String(numFmts.findAll(MAIN, "numFmt").length));
    return newId;
  }

  private get cellXfsElOrThrow(): XmlElement {
    const el = this.cellXfsEl;
    if (el === null) throw new InvalidXmlError("stylesheet has no <cellXfs>");
    return el;
  }

  // -- style-table reads (per style index `s`) --------------------------

  private xfAt(s: number): XmlElement | null {
    return this.cellXfsElOrThrow.findAll(MAIN, "xf")[s] ?? null;
  }

  private tableChildAt(container: string, id: number): XmlElement | null {
    const el = this.el.find(MAIN, container);
    if (el === null) return null;
    return el.childElements[id] ?? null;
  }

  /** Font applied by cell style index `s`. */
  fontForStyle(s: number): Font {
    const xf = this.xfAt(s);
    const fontId = xf === null ? 0 : Number(xf.getAttr("fontId") ?? "0");
    const fontEl = this.tableChildAt("fonts", fontId);
    return fontEl === null ? new Font() : Font.fromElement(fontEl);
  }

  /** Fill applied by cell style index `s`. */
  fillForStyle(s: number): Fill {
    const xf = this.xfAt(s);
    const fillId = xf === null ? 0 : Number(xf.getAttr("fillId") ?? "0");
    const fillEl = this.tableChildAt("fills", fillId);
    return fillEl === null ? new PatternFill() : fillFromElement(fillEl);
  }

  /** Border applied by cell style index `s`. */
  borderForStyle(s: number): Border {
    const xf = this.xfAt(s);
    const borderId = xf === null ? 0 : Number(xf.getAttr("borderId") ?? "0");
    const borderEl = this.tableChildAt("borders", borderId);
    return borderEl === null ? new Border() : Border.fromElement(borderEl);
  }

  /** Alignment applied by cell style index `s`. */
  alignmentForStyle(s: number): Alignment {
    return Alignment.fromElement(this.xfAt(s)?.find(MAIN, "alignment") ?? null);
  }

  /** Protection applied by cell style index `s`. */
  protectionForStyle(s: number): Protection {
    return Protection.fromElement(this.xfAt(s)?.find(MAIN, "protection") ?? null);
  }

  // -- style-table find-or-add (never bloats a table) -------------------

  private getOrAddTableEntry(
    container: string,
    build: (el: XmlElement) => void,
    parse: (el: XmlElement) => { key(): string },
    targetKey: string,
  ): number {
    let tableEl = this.el.find(MAIN, container);
    if (tableEl === null) {
      tableEl = createElement(`main:${container}` as NsTag, this.el);
      this.insertOrdered(tableEl, container);
    }
    const entries = tableEl.childElements;
    for (let i = 0; i < entries.length; i++) {
      if (parse(entries[i]!).key() === targetKey) return i;
    }
    const singular = container.replace(/s$/, ""); // fonts→font, fills→fill, borders→border
    const el = createElement(`main:${singular}` as NsTag, tableEl);
    build(el);
    tableEl.appendChild(el);
    tableEl.setAttr("count", String(entries.length + 1));
    return entries.length;
  }

  getOrAddFont(font: Font): number {
    return this.getOrAddTableEntry("fonts", (el) => font.writeInto(el), Font.fromElement, font.key());
  }

  getOrAddFill(fill: Fill): number {
    return this.getOrAddTableEntry("fills", (el) => fill.writeInto(el), fillFromElement, fill.key());
  }

  getOrAddBorder(border: Border): number {
    return this.getOrAddTableEntry(
      "borders",
      (el) => border.writeInto(el),
      Border.fromElement,
      border.key(),
    );
  }

  /**
   * Find-or-add a cellXfs entry equal to `baseIndex`'s xf with `overrides`
   * applied. Reuses an identical existing xf so the table never bloats; only
   * the dimensions named in `overrides` change (others inherit from the base).
   */
  getOrAddXf(baseIndex: number, overrides: XfOverrides): number {
    const cellXfs = this.cellXfsElOrThrow;
    const xfs = cellXfs.findAll(MAIN, "xf");
    const base = xfs[baseIndex] ?? xfs[0] ?? null;
    const target = base === null ? createElement("main:xf", cellXfs) : base.cloneNode();

    const applyAttr = (name: string, id: number | undefined, applyName: string): void => {
      if (id === undefined) return;
      target.setAttr(name, String(id));
      target.setAttr(applyName, "1");
    };
    applyAttr("numFmtId", overrides.numFmtId, "applyNumberFormat");
    applyAttr("fontId", overrides.fontId, "applyFont");
    applyAttr("fillId", overrides.fillId, "applyFill");
    applyAttr("borderId", overrides.borderId, "applyBorder");

    if (overrides.alignment !== undefined) {
      setXfSubElement(target, "alignment", overrides.alignment.isEmpty ? null : overrides.alignment);
      if (overrides.alignment.isEmpty) target.removeAttr("applyAlignment");
      else target.setAttr("applyAlignment", "1");
    }
    if (overrides.protection !== undefined) {
      setXfSubElement(target, "protection", overrides.protection.isEmpty ? null : overrides.protection);
      if (overrides.protection.isEmpty) target.removeAttr("applyProtection");
      else target.setAttr("applyProtection", "1");
    }

    const key = xfKey(target);
    for (let i = 0; i < xfs.length; i++) {
      if (xfKey(xfs[i]!) === key) return i;
    }
    cellXfs.appendChild(target);
    cellXfs.setAttr("count", String(xfs.length + 1));
    return xfs.length;
  }

  /** Back-compat alias used by the Date write path (numFmt only). */
  getOrAddXfWithNumFmt(baseIndex: number, numFmtId: number): number {
    return this.getOrAddXf(baseIndex, { numFmtId });
  }

  // -- named styles (cellStyles + cellStyleXfs) -------------------------

  /** Names of the workbook's named cell styles ("Normal", "Title", …). */
  get namedStyleNames(): string[] {
    const cs = this.el.find(MAIN, "cellStyles");
    if (cs === null) return [];
    return cs.findAll(MAIN, "cellStyle").map((e) => e.getAttr("name") ?? "");
  }

  private cellStyleXfIdForName(name: string): number | null {
    const cs = this.el.find(MAIN, "cellStyles");
    if (cs === null) return null;
    for (const e of cs.findAll(MAIN, "cellStyle")) {
      if (e.getAttr("name") === name) return Number(e.getAttr("xfId") ?? "0");
    }
    return null;
  }

  /** Name of the named style referenced by cell style index `s` (xf/@xfId → a
   * cellStyle), defaulting to "Normal". */
  styleNameForStyle(s: number): string {
    const xf = this.xfAt(s);
    const xfId = xf === null ? 0 : Number(xf.getAttr("xfId") ?? "0");
    const cs = this.el.find(MAIN, "cellStyles");
    if (cs !== null) {
      for (const e of cs.findAll(MAIN, "cellStyle")) {
        if (Number(e.getAttr("xfId") ?? "0") === xfId) return e.getAttr("name") ?? "Normal";
      }
    }
    return "Normal";
  }

  /** Find-or-add a differential format (`<dxf>`) carrying a solid fill, used by
   * conditional-formatting rules. Returns the dxfId. dxf fills conventionally
   * carry their color in `<bgColor>`. */
  getOrAddDxfFill(fill: PatternFill): number {
    let dxfs = this.el.find(MAIN, "dxfs");
    if (dxfs === null) {
      dxfs = createElement("main:dxfs", this.el);
      this.insertOrdered(dxfs, "dxfs");
    }
    // Build attached-as-we-go so createElement resolves the default namespace.
    const dxf = createElement("main:dxf", dxfs);
    dxfs.appendChild(dxf);
    const fillEl = createElement("main:fill", dxf);
    dxf.appendChild(fillEl);
    const pf = createElement("main:patternFill", fillEl);
    fillEl.appendChild(pf);
    const color = fill.fgColor ?? fill.bgColor;
    if (color !== undefined) {
      const bg = createElement("main:bgColor", pf);
      pf.appendChild(bg);
      color.writeInto(bg);
    }
    const count = dxfs.findAll(MAIN, "dxf").length;
    dxfs.setAttr("count", String(count));
    return count - 1;
  }

  /** Apply an existing named style to a cell, returning a (reused/new) cellXfs
   * index whose xf mirrors the named style's cellStyleXf and points at it. */
  applyNamedStyle(name: string): number {
    const xfId = this.cellStyleXfIdForName(name);
    if (xfId === null) throw new InvalidXmlError(`no named style "${name}"`);
    const styleXfs = this.el.find(MAIN, "cellStyleXfs");
    const styleXf = styleXfs?.childElements[xfId] ?? null;
    const cellXfs = this.cellXfsElOrThrow;
    const xfs = cellXfs.findAll(MAIN, "xf");

    const target = styleXf === null ? createElement("main:xf", cellXfs) : styleXf.cloneNode();
    target.setAttr("xfId", String(xfId));

    const key = xfKey(target);
    for (let i = 0; i < xfs.length; i++) if (xfKey(xfs[i]!) === key) return i;
    cellXfs.appendChild(target);
    cellXfs.setAttr("count", String(xfs.length + 1));
    return xfs.length;
  }
}

/** Replace an xf's `<alignment>` / `<protection>` child (child order:
 * alignment, protection, extLst). A null value removes it. */
function setXfSubElement(
  xf: XmlElement,
  name: "alignment" | "protection",
  value: Alignment | Protection | null,
): void {
  const existing = xf.find(MAIN, name);
  if (existing !== null) xf.removeChild(existing);
  if (value === null) return;
  const el = createElement(`main:${name}` as NsTag, xf);
  value.writeInto(el);
  // insert before the first successor (protection/extLst for alignment; extLst for protection)
  const successors = name === "alignment" ? ["protection", "extLst"] : ["extLst"];
  for (const succ of successors) {
    const ref = xf.find(MAIN, succ);
    if (ref !== null) {
      xf.insertBefore(el, ref);
      return;
    }
  }
  xf.appendChild(el);
}

/** Canonical identity of an `<xf>`: sorted attrs + child structure, so that
 * two xfs that would render identically dedupe to one cellXfs entry. */
function xfKey(xf: XmlElement): string {
  const attrs = xf.attrs
    .map((a) => `${a.name}=${a.rawValue}`)
    .sort()
    .join("|");
  const kids = xf.childElements
    .map((c) => `${c.name}:${c.attrs.map((a) => `${a.name}=${a.rawValue}`).sort().join(",")}`)
    .join(";");
  return `${attrs}#${kids}`;
}
