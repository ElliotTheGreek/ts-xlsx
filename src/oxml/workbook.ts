/**
 * Workbook element wrappers — port of openpyxl workbook parsing.
 * `<workbook>` holds `<sheets>` (one `<sheet>` per worksheet, referencing its
 * part by r:id), `<definedNames>`, `<calcPr>` (fullCalcOnLoad), and
 * `<workbookPr>` (date1904 flag).
 */
import { XmlElement } from "../xml/dom.js";
import { OxmlWrapper, createElement } from "./base.js";
import { nsmap, NsTag } from "./ns.js";
import { XsdBoolean } from "./simpletypes.js";

const MAIN = nsmap.main;
const R = nsmap.r;

/** Child order under `<workbook>` (ECMA-376 §18.2.27), for insertion. */
const WORKBOOK_CHILDREN: readonly NsTag[] = [
  "main:fileVersion",
  "main:fileSharing",
  "main:workbookPr",
  "main:workbookProtection",
  "main:bookViews",
  "main:sheets",
  "main:functionGroups",
  "main:externalReferences",
  "main:definedNames",
  "main:calcPr",
] as const;

/** `<sheet>` — one worksheet entry in `<sheets>`. */
export class CT_Sheet extends OxmlWrapper {
  get name(): string {
    return this.el.getAttr("name") ?? "";
  }
  set name(v: string) {
    this.el.setAttr("name", v);
  }

  get sheetId(): number {
    const v = this.el.getAttr("sheetId");
    return v === null ? 0 : Number(v);
  }

  /** "visible" | "hidden" | "veryHidden" (default "visible"). */
  get state(): string {
    return this.el.getAttr("state") ?? "visible";
  }
  set state(v: string) {
    if (v === "visible") this.el.removeAttr("state");
    else this.el.setAttr("state", v);
  }

  /** The r:id linking to the worksheet part. */
  get rId(): string | null {
    return this.el.getAttrNS(R, "id");
  }
}

/** `<workbook>` root. */
export class CT_Workbook extends OxmlWrapper {
  private get sheetsEl(): XmlElement | null {
    return this.el.find(MAIN, "sheets");
  }

  get sheetLst(): CT_Sheet[] {
    const sheets = this.sheetsEl;
    if (sheets === null) return [];
    return sheets.findAll(MAIN, "sheet").map((e) => new CT_Sheet(e));
  }

  /** Ensure `<sheets>` exists (M2/M4 create path). */
  getOrAddSheets(): XmlElement {
    return this.getOrAdd("main:sheets", successorsAfter("main:sheets"));
  }

  /** `<calcPr>` element, or null. */
  get calcPr(): XmlElement | null {
    return this.el.find(MAIN, "calcPr");
  }

  getOrAddCalcPr(): XmlElement {
    return this.getOrAdd("main:calcPr", successorsAfter("main:calcPr"));
  }

  /** Whether the workbook uses the 1904 date system (`workbookPr/@date1904`). */
  get date1904(): boolean {
    const pr = this.el.find(MAIN, "workbookPr");
    if (pr === null) return false;
    const v = pr.getAttr("date1904");
    return v === null ? false : XsdBoolean.fromXml(v);
  }

  /** Set (or clear) calcPr/@fullCalcOnLoad. */
  setFullCalcOnLoad(on: boolean): void {
    const calcPr = this.getOrAddCalcPr();
    if (on) calcPr.setAttr("fullCalcOnLoad", "1");
    else calcPr.removeAttr("fullCalcOnLoad");
  }

  get definedNamesEl(): XmlElement | null {
    return this.el.find(MAIN, "definedNames");
  }

  /** Defined names: name → refers-to expression. */
  get definedNamesMap(): Map<string, string> {
    const out = new Map<string, string>();
    const dn = this.definedNamesEl;
    if (dn === null) return out;
    for (const e of dn.findAll(MAIN, "definedName")) {
      out.set(e.getAttr("name") ?? "", e.text);
    }
    return out;
  }

  setDefinedName(name: string, refersTo: string): void {
    const dn = this.getOrAdd("main:definedNames", successorsAfter("main:definedNames"));
    for (const e of dn.findAll(MAIN, "definedName")) {
      if (e.getAttr("name") === name) {
        e.setText(refersTo);
        return;
      }
    }
    const el = createElement("main:definedName", dn);
    el.setAttr("name", name);
    dn.appendChild(el);
    el.setText(refersTo);
  }

  removeDefinedName(name: string): void {
    const dn = this.definedNamesEl;
    if (dn === null) return;
    for (const e of dn.findAll(MAIN, "definedName")) {
      if (e.getAttr("name") === name) {
        dn.removeChild(e);
        break;
      }
    }
    if (dn.findAll(MAIN, "definedName").length === 0) this.el.removeChild(dn);
  }

  // -- sheet element management (M4) ------------------------------------

  /** Next free sheetId (max existing + 1, min 1). */
  nextSheetId(): number {
    let max = 0;
    for (const s of this.sheetLst) max = Math.max(max, s.sheetId);
    return max + 1;
  }

  /** Add a `<sheet>` referencing a worksheet part by rId, at `index` (or end). */
  addSheet(name: string, sheetId: number, rId: string, index?: number): CT_Sheet {
    const sheets = this.getOrAddSheets();
    const el = createElement("main:sheet", sheets);
    el.setAttr("name", name);
    el.setAttr("sheetId", String(sheetId));
    el.setAttr("xmlns:r", R);
    el.setAttr("r:id", rId);
    const existing = sheets.findAll(MAIN, "sheet");
    if (index === undefined || index >= existing.length) {
      sheets.appendChild(el);
    } else {
      sheets.insertBefore(el, existing[index]!);
    }
    return new CT_Sheet(el);
  }

  /** Remove the `<sheet>` with `name`; returns its rId (or null if absent). */
  removeSheet(name: string): string | null {
    const sheets = this.el.find(MAIN, "sheets");
    if (sheets === null) return null;
    for (const el of sheets.findAll(MAIN, "sheet")) {
      if (el.getAttr("name") === name) {
        const rId = el.getAttrNS(R, "id");
        sheets.removeChild(el);
        return rId;
      }
    }
    return null;
  }

  /** Move the `<sheet>` named `name` to absolute `index`. */
  moveSheet(name: string, index: number): void {
    const sheets = this.el.find(MAIN, "sheets");
    if (sheets === null) return;
    const els = sheets.findAll(MAIN, "sheet");
    const el = els.find((e) => e.getAttr("name") === name);
    if (el === undefined) return;
    sheets.removeChild(el);
    const after = sheets.findAll(MAIN, "sheet");
    const clamped = Math.max(0, Math.min(index, after.length));
    if (clamped >= after.length) sheets.appendChild(el);
    else sheets.insertBefore(el, after[clamped]!);
  }
}

/** Successor tags (those that must come after `tag`) within `<workbook>`. */
export function successorsAfter(tag: NsTag): NsTag[] {
  const i = WORKBOOK_CHILDREN.indexOf(tag);
  return i === -1 ? [] : WORKBOOK_CHILDREN.slice(i + 1);
}
