/**
 * Worksheet element wrappers — port of openpyxl worksheet parsing/writing.
 *
 * `<worksheet>` → `<sheetData>` → `<row r="N">` → `<c r="A1" t="…" s="…">` with
 * a `<v>` value, an optional `<f>` formula, or an `<is>` inline string. Rows are
 * kept sorted by `@r` and cells within a row sorted by column, matching Excel's
 * canonical ordering so a re-emitted sheet stays clean.
 */
import { XmlElement } from "../xml/dom.js";
import { OxmlWrapper, createElement } from "./base.js";
import { nsmap, NsTag } from "./ns.js";
import {
  cellRefToRowCol,
  columnIndexFromString,
  coordinateFromString,
  getColumnLetter,
} from "../util.js";

const MAIN = nsmap.main;

/** Child order under `<c>` (ECMA-376 §18.3.1.4). */
const CELL_CHILDREN: readonly NsTag[] = ["main:f", "main:v", "main:is"] as const;

function cellSuccessors(tag: NsTag): NsTag[] {
  const i = CELL_CHILDREN.indexOf(tag);
  return i === -1 ? [] : CELL_CHILDREN.slice(i + 1);
}

/** `<c>` cell element. */
export class CT_Cell extends OxmlWrapper {
  /** The A1 reference (`@r`). */
  get ref(): string {
    return this.el.getAttr("r") ?? "";
  }
  set ref(v: string) {
    this.el.setAttr("r", v);
  }

  /** Data type (`@t`), defaulting to "n" (number) when absent. */
  get typeAttr(): string {
    return this.el.getAttr("t") ?? "n";
  }
  /** Set `@t`; "n" removes the attribute (its schema default). */
  set typeAttr(t: string) {
    if (t === "n") this.el.removeAttr("t");
    else this.el.setAttr("t", t);
  }

  /** Style index (`@s`), 0 when absent. */
  get styleIndex(): number {
    const v = this.el.getAttr("s");
    return v === null ? 0 : Number(v);
  }
  set styleIndex(i: number) {
    if (i === 0) this.el.removeAttr("s");
    else this.el.setAttr("s", String(i));
  }

  get vEl(): XmlElement | null {
    return this.el.find(MAIN, "v");
  }
  get fEl(): XmlElement | null {
    return this.el.find(MAIN, "f");
  }
  get isEl(): XmlElement | null {
    return this.el.find(MAIN, "is");
  }

  /** Raw `<v>` text, or null. */
  get vText(): string | null {
    return this.vEl?.text ?? null;
  }

  /** Concatenated text of an `<is>` inline string (rich runs joined). */
  get inlineText(): string | null {
    const is = this.isEl;
    if (is === null) return null;
    const directT = is.find(MAIN, "t");
    if (directT !== null) return directT.text;
    let out = "";
    for (const r of is.findAll(MAIN, "r")) {
      const t = r.find(MAIN, "t");
      if (t !== null) out += t.text;
    }
    return out;
  }

  /** Formula string (without leading "="), or null when the cell has no `<f>`. */
  get formula(): string | null {
    return this.fEl?.text ?? null;
  }

  // -- write helpers (M2) ------------------------------------------------

  getOrAddV(): XmlElement {
    let v = this.vEl;
    if (v === null) {
      v = createElement("main:v", this.el);
      this.insertBeforeSuccessors(v, cellSuccessors("main:v"));
    }
    return v;
  }

  removeChildTag(tag: "f" | "v" | "is"): void {
    const el = this.el.find(MAIN, tag);
    if (el !== null) this.el.removeChild(el);
  }

  /** Clear every value child (`<f>`, `<v>`, `<is>`). */
  clearValue(): void {
    for (const tag of ["f", "v", "is"] as const) this.removeChildTag(tag);
  }

  /** Set `<f>` formula text (creating it before `<v>`). */
  setFormula(text: string): XmlElement {
    let f = this.fEl;
    if (f === null) {
      f = createElement("main:f", this.el);
      this.insertBeforeSuccessors(f, cellSuccessors("main:f"));
    }
    f.setText(text);
    return f;
  }

  /** Replace content with an inline string `<is><t>…</t></is>`. */
  setInlineString(text: string): void {
    this.clearValue();
    const is = createElement("main:is", this.el);
    this.insertBeforeSuccessors(is, cellSuccessors("main:is"));
    const t = createElement("main:t", is);
    is.appendChild(t);
    t.setText(text);
    if (text.trim().length !== text.length) t.setAttr("xml:space", "preserve");
  }

  private insertBeforeSuccessors(child: XmlElement, successors: readonly NsTag[]): void {
    for (const succ of successors) {
      const ref = this.el.find(MAIN, succ.slice(succ.indexOf(":") + 1));
      if (ref !== null) {
        this.el.insertBefore(child, ref);
        return;
      }
    }
    this.el.appendChild(child);
  }
}

/** `<row>` element. */
export class CT_Row extends OxmlWrapper {
  /** Row number (`@r`). */
  get r(): number {
    const v = this.el.getAttr("r");
    return v === null ? 0 : Number(v);
  }
  set r(v: number) {
    this.el.setAttr("r", String(v));
  }

  get cellLst(): CT_Cell[] {
    return this.el.findAll(MAIN, "c").map((e) => new CT_Cell(e));
  }

  /** Find the `<c>` with this A1 ref, or null. */
  findCell(ref: string): CT_Cell | null {
    for (const c of this.el.findAll(MAIN, "c")) {
      if (c.getAttr("r") === ref) return new CT_Cell(c);
    }
    return null;
  }

  /** Find-or-create the `<c>` at A1 `ref`, kept sorted by column. */
  getOrAddCell(ref: string): CT_Cell {
    const existing = this.findCell(ref);
    if (existing !== null) return existing;
    const col = columnIndexFromString(coordinateFromString(ref).column);
    const c = createElement("main:c", this.el);
    c.setAttr("r", ref);
    // insert before the first cell whose column exceeds `col`
    let inserted = false;
    for (const cur of this.el.findAll(MAIN, "c")) {
      const curCol = cellRefToRowCol(cur.getAttr("r") ?? "A1").col;
      if (curCol > col) {
        this.el.insertBefore(c, cur);
        inserted = true;
        break;
      }
    }
    if (!inserted) this.el.appendChild(c);
    return new CT_Cell(c);
  }
}

/** `<sheetData>` element. */
export class CT_SheetData extends OxmlWrapper {
  get rowLst(): CT_Row[] {
    return this.el.findAll(MAIN, "row").map((e) => new CT_Row(e));
  }

  findRow(r: number): CT_Row | null {
    for (const row of this.el.findAll(MAIN, "row")) {
      if (Number(row.getAttr("r")) === r) return new CT_Row(row);
    }
    return null;
  }

  /** Find-or-create the `<row r="r">`, kept sorted by row number. */
  getOrAddRow(r: number): CT_Row {
    const existing = this.findRow(r);
    if (existing !== null) return existing;
    const row = createElement("main:row", this.el);
    row.setAttr("r", String(r));
    let inserted = false;
    for (const cur of this.el.findAll(MAIN, "row")) {
      if (Number(cur.getAttr("r")) > r) {
        this.el.insertBefore(row, cur);
        inserted = true;
        break;
      }
    }
    if (!inserted) this.el.appendChild(row);
    return new CT_Row(row);
  }
}

/** Child order under `<worksheet>` (ECMA-376 §18.3.1.99), for insertion. */
const WORKSHEET_CHILDREN = [
  "sheetPr",
  "dimension",
  "sheetViews",
  "sheetFormatPr",
  "cols",
  "sheetData",
  "sheetCalcPr",
  "sheetProtection",
  "protectedRanges",
  "scenarios",
  "autoFilter",
  "sortState",
  "dataConsolidate",
  "customSheetViews",
  "mergeCells",
  "phoneticPr",
  "conditionalFormatting",
  "dataValidations",
  "hyperlinks",
  "printOptions",
  "pageMargins",
  "pageSetup",
  "headerFooter",
  "rowBreaks",
  "colBreaks",
  "customProperties",
  "cellWatches",
  "ignoredErrors",
  "smartTags",
  "drawing",
  "legacyDrawing",
  "tableParts",
  "extLst",
] as const;

/** `<worksheet>` root. */
export class CT_Worksheet extends OxmlWrapper {
  get sheetData(): CT_SheetData {
    return new CT_SheetData(this.oneAndOnlyOne("main:sheetData"));
  }

  /** Find-or-create a direct child, inserted in schema order. */
  getOrAddOrdered(localName: (typeof WORKSHEET_CHILDREN)[number]): XmlElement {
    const existing = this.el.find(MAIN, localName);
    if (existing !== null) return existing;
    const el = createElement(`main:${localName}` as NsTag, this.el);
    const i = WORKSHEET_CHILDREN.indexOf(localName);
    for (const succ of WORKSHEET_CHILDREN.slice(i + 1)) {
      const ref = this.el.find(MAIN, succ);
      if (ref !== null) {
        this.el.insertBefore(el, ref);
        return el;
      }
    }
    this.el.appendChild(el);
    return el;
  }

  /** `<dimension>` @ref, or null. */
  get dimensionRef(): string | null {
    return this.el.find(MAIN, "dimension")?.getAttr("ref") ?? null;
  }
  set dimensionRef(ref: string) {
    this.getOrAddOrdered("dimension").setAttr("ref", ref);
  }

  // -- merged ranges ----------------------------------------------------

  get mergedRefs(): string[] {
    const mc = this.el.find(MAIN, "mergeCells");
    if (mc === null) return [];
    return mc.findAll(MAIN, "mergeCell").map((e) => e.getAttr("ref") ?? "");
  }

  addMerge(ref: string): void {
    const mc = this.getOrAddOrdered("mergeCells");
    for (const e of mc.findAll(MAIN, "mergeCell")) if (e.getAttr("ref") === ref) return;
    const cell = createElement("main:mergeCell", mc);
    cell.setAttr("ref", ref);
    mc.appendChild(cell);
    mc.setAttr("count", String(mc.findAll(MAIN, "mergeCell").length));
  }

  removeMerge(ref: string): void {
    const mc = this.el.find(MAIN, "mergeCells");
    if (mc === null) return;
    for (const e of mc.findAll(MAIN, "mergeCell")) {
      if (e.getAttr("ref") === ref) {
        mc.removeChild(e);
        break;
      }
    }
    const remaining = mc.findAll(MAIN, "mergeCell");
    if (remaining.length === 0) this.el.removeChild(mc);
    else mc.setAttr("count", String(remaining.length));
  }

  /** Rewrite the ref list wholesale (used by row/col shifting). */
  replaceMerges(refs: string[]): void {
    const mc = this.el.find(MAIN, "mergeCells");
    if (mc === null) {
      if (refs.length === 0) return;
    } else {
      this.el.removeChild(mc);
    }
    if (refs.length === 0) return;
    const el = this.getOrAddOrdered("mergeCells");
    for (const ref of refs) {
      const cell = createElement("main:mergeCell", el);
      cell.setAttr("ref", ref);
      el.appendChild(cell);
    }
    el.setAttr("count", String(refs.length));
  }

  // -- columns ----------------------------------------------------------

  /** Find-or-create a single-column `<col min=idx max=idx>` entry. */
  getOrAddCol(idx: number): XmlElement {
    const cols = this.getOrAddOrdered("cols");
    for (const c of cols.findAll(MAIN, "col")) {
      if (Number(c.getAttr("min")) === idx && Number(c.getAttr("max")) === idx) return c;
    }
    // reuse a spanning col that covers idx (edit affects its whole span)
    for (const c of cols.findAll(MAIN, "col")) {
      const min = Number(c.getAttr("min"));
      const max = Number(c.getAttr("max"));
      if (min <= idx && idx <= max && min === max) return c;
    }
    const col = createElement("main:col", cols);
    col.setAttr("min", String(idx));
    col.setAttr("max", String(idx));
    // insert keeping ascending min order
    let inserted = false;
    for (const c of cols.findAll(MAIN, "col")) {
      if (Number(c.getAttr("min")) > idx) {
        cols.insertBefore(col, c);
        inserted = true;
        break;
      }
    }
    if (!inserted) cols.appendChild(col);
    return col;
  }

  findCol(idx: number): XmlElement | null {
    const cols = this.el.find(MAIN, "cols");
    if (cols === null) return null;
    for (const c of cols.findAll(MAIN, "col")) {
      if (Number(c.getAttr("min")) <= idx && idx <= Number(c.getAttr("max"))) return c;
    }
    return null;
  }

  get colEls(): XmlElement[] {
    return this.el.find(MAIN, "cols")?.findAll(MAIN, "col") ?? [];
  }

  // -- sheet view / freeze panes ---------------------------------------

  private getOrAddSheetView(): XmlElement {
    const views = this.getOrAddOrdered("sheetViews");
    let view = views.find(MAIN, "sheetView");
    if (view === null) {
      view = createElement("main:sheetView", views);
      view.setAttr("workbookViewId", "0");
      views.appendChild(view);
    }
    return view;
  }

  get paneTopLeft(): string | null {
    const pane = this.el.find(MAIN, "sheetViews")?.find(MAIN, "sheetView")?.find(MAIN, "pane");
    return pane?.getAttr("topLeftCell") ?? null;
  }

  setFreezePane(xSplit: number, ySplit: number, topLeft: string, activePane: string): void {
    const view = this.getOrAddSheetView();
    const existing = view.find(MAIN, "pane");
    if (existing !== null) view.removeChild(existing);
    const pane = createElement("main:pane", view);
    if (xSplit > 0) pane.setAttr("xSplit", String(xSplit));
    if (ySplit > 0) pane.setAttr("ySplit", String(ySplit));
    pane.setAttr("topLeftCell", topLeft);
    pane.setAttr("activePane", activePane);
    pane.setAttr("state", "frozen");
    // pane must be the first child of sheetView
    const first = view.childElements[0] ?? null;
    if (first !== null) view.insertBefore(pane, first);
    else view.appendChild(pane);
  }

  clearFreezePane(): void {
    const view = this.el.find(MAIN, "sheetViews")?.find(MAIN, "sheetView");
    const pane = view?.find(MAIN, "pane");
    if (view && pane) view.removeChild(pane);
  }

  // -- autofilter -------------------------------------------------------

  get autoFilterRef(): string | null {
    return this.el.find(MAIN, "autoFilter")?.getAttr("ref") ?? null;
  }
  setAutoFilter(ref: string | null): void {
    if (ref === null) {
      const af = this.el.find(MAIN, "autoFilter");
      if (af !== null) this.el.removeChild(af);
      return;
    }
    this.getOrAddOrdered("autoFilter").setAttr("ref", ref);
  }

  // -- tab color (sheetPr/tabColor) ------------------------------------

  get tabColorRgb(): string | null {
    return this.el.find(MAIN, "sheetPr")?.find(MAIN, "tabColor")?.getAttr("rgb") ?? null;
  }
  setTabColorRgb(rgb: string | null): void {
    const sheetPr = this.getOrAddOrdered("sheetPr");
    const existing = sheetPr.find(MAIN, "tabColor");
    if (rgb === null) {
      if (existing !== null) sheetPr.removeChild(existing);
      return;
    }
    const tab = existing ?? createElement("main:tabColor", sheetPr);
    tab.setAttr("rgb", rgb);
    if (existing === null) {
      // tabColor is the first child of sheetPr
      const first = sheetPr.childElements[0] ?? null;
      if (first !== null) sheetPr.insertBefore(tab, first);
      else sheetPr.appendChild(tab);
    }
  }

  // -- row / column insert & delete (cell-ref shifting) -----------------
  // Formulas are NOT rewritten on shift (openpyxl behaves the same); only the
  // structural @r references of cells, rows, cols and merges move.

  private shiftRowRefs(rowEl: XmlElement, delta: number): void {
    const newR = Number(rowEl.getAttr("r")) + delta;
    rowEl.setAttr("r", String(newR));
    for (const c of rowEl.findAll(MAIN, "c")) {
      const { col } = cellRefToRowCol(c.getAttr("r") ?? "A1");
      c.setAttr("r", `${getColumnLetter(col)}${newR}`);
    }
  }

  insertRows(fromRow: number, count: number): void {
    const sheetData = this.oneAndOnlyOne("main:sheetData");
    for (const rowEl of sheetData.findAll(MAIN, "row")) {
      if (Number(rowEl.getAttr("r")) >= fromRow) this.shiftRowRefs(rowEl, count);
    }
  }

  deleteRows(fromRow: number, count: number): void {
    const sheetData = this.oneAndOnlyOne("main:sheetData");
    for (const rowEl of sheetData.findAll(MAIN, "row")) {
      const r = Number(rowEl.getAttr("r"));
      if (r >= fromRow && r < fromRow + count) sheetData.removeChild(rowEl);
    }
    for (const rowEl of sheetData.findAll(MAIN, "row")) {
      if (Number(rowEl.getAttr("r")) >= fromRow + count) this.shiftRowRefs(rowEl, -count);
    }
  }

  insertCols(fromCol: number, count: number): void {
    const sheetData = this.oneAndOnlyOne("main:sheetData");
    for (const rowEl of sheetData.findAll(MAIN, "row")) {
      const r = Number(rowEl.getAttr("r"));
      // shift right-to-left so ascending order is preserved as we rename
      const cells = rowEl.findAll(MAIN, "c").reverse();
      for (const c of cells) {
        const { col } = cellRefToRowCol(c.getAttr("r") ?? "A1");
        if (col >= fromCol) c.setAttr("r", `${getColumnLetter(col + count)}${r}`);
      }
    }
    this.shiftColDims(fromCol, count);
  }

  deleteCols(fromCol: number, count: number): void {
    const sheetData = this.oneAndOnlyOne("main:sheetData");
    for (const rowEl of sheetData.findAll(MAIN, "row")) {
      const r = Number(rowEl.getAttr("r"));
      for (const c of rowEl.findAll(MAIN, "c")) {
        const { col } = cellRefToRowCol(c.getAttr("r") ?? "A1");
        if (col >= fromCol && col < fromCol + count) rowEl.removeChild(c);
        else if (col >= fromCol + count) c.setAttr("r", `${getColumnLetter(col - count)}${r}`);
      }
    }
    this.shiftColDims(fromCol, -count);
  }

  private shiftColDims(fromCol: number, delta: number): void {
    const cols = this.el.find(MAIN, "cols");
    if (cols === null) return;
    for (const c of cols.findAll(MAIN, "col")) {
      const min = Number(c.getAttr("min"));
      const max = Number(c.getAttr("max"));
      if (delta < 0 && min >= fromCol && max < fromCol - delta) {
        cols.removeChild(c);
        continue;
      }
      if (min >= fromCol) c.setAttr("min", String(Math.max(fromCol, min + delta)));
      if (max >= fromCol) c.setAttr("max", String(max + delta));
    }
    if (cols.findAll(MAIN, "col").length === 0) this.el.removeChild(cols);
  }

  // -- conditional formatting (M5) --------------------------------------

  get conditionalFormattingEls(): XmlElement[] {
    return this.el.findAll(MAIN, "conditionalFormatting");
  }

  addConditionalFormatting(sqref: string): XmlElement {
    const cf = createElement("main:conditionalFormatting", this.el);
    cf.setAttr("sqref", sqref);
    // conditionalFormatting sits after mergeCells/phoneticPr, before dataValidations
    const i = WORKSHEET_CHILDREN.indexOf("conditionalFormatting");
    for (const succ of WORKSHEET_CHILDREN.slice(i + 1)) {
      const ref = this.el.find(MAIN, succ);
      if (ref !== null) {
        this.el.insertBefore(cf, ref);
        return cf;
      }
    }
    this.el.appendChild(cf);
    return cf;
  }

  /** Highest cfRule @priority currently in use (0 when none). */
  maxCfPriority(): number {
    let max = 0;
    for (const cf of this.conditionalFormattingEls) {
      for (const rule of cf.findAll(MAIN, "cfRule")) {
        max = Math.max(max, Number(rule.getAttr("priority") ?? "0"));
      }
    }
    return max;
  }

  // -- data validation (M5) ---------------------------------------------

  get dataValidationEls(): XmlElement[] {
    return this.el.find(MAIN, "dataValidations")?.findAll(MAIN, "dataValidation") ?? [];
  }

  addDataValidation(): XmlElement {
    const dvs = this.getOrAddOrdered("dataValidations");
    const dv = createElement("main:dataValidation", dvs);
    dvs.appendChild(dv);
    dvs.setAttr("count", String(dvs.findAll(MAIN, "dataValidation").length));
    return dv;
  }

  // -- hyperlinks (M5) --------------------------------------------------

  get hyperlinkEls(): XmlElement[] {
    return this.el.find(MAIN, "hyperlinks")?.findAll(MAIN, "hyperlink") ?? [];
  }

  addHyperlinkEl(ref: string): XmlElement {
    const hs = this.getOrAddOrdered("hyperlinks");
    for (const h of hs.findAll(MAIN, "hyperlink")) {
      if (h.getAttr("ref") === ref) return h;
    }
    const h = createElement("main:hyperlink", hs);
    h.setAttr("ref", ref);
    hs.appendChild(h);
    return h;
  }

  removeHyperlinkEl(ref: string): string | null {
    const hs = this.el.find(MAIN, "hyperlinks");
    if (hs === null) return null;
    for (const h of hs.findAll(MAIN, "hyperlink")) {
      if (h.getAttr("ref") === ref) {
        const rId = h.getAttrNS(nsmap.r, "id");
        hs.removeChild(h);
        if (hs.findAll(MAIN, "hyperlink").length === 0) this.el.removeChild(hs);
        return rId;
      }
    }
    return null;
  }

  // -- drawing reference (M6) -------------------------------------------

  get drawingRId(): string | null {
    return this.el.find(MAIN, "drawing")?.getAttrNS(nsmap.r, "id") ?? null;
  }

  setDrawingRel(rId: string): void {
    let drawing = this.el.find(MAIN, "drawing");
    if (drawing === null) drawing = this.getOrAddOrdered("drawing");
    drawing.setAttr("xmlns:r", nsmap.r);
    drawing.setAttr("r:id", rId);
  }
}
