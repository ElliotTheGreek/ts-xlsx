/**
 * Worksheet — port of openpyxl/worksheet/worksheet.py (read surface).
 *
 * Cell access returns lazy Cell proxies and never materializes `<c>` elements
 * on read. maxRow/maxColumn/dimensions are computed from the populated cells,
 * matching openpyxl's calculate_dimension.
 */
import { Cell } from "./cell.js";
import type { CellValue } from "./cell.js";
import {
  cellRefToRowCol,
  columnIndexFromString,
  getColumnLetter,
  rowColToCellRef,
} from "./util.js";
import { readFileSync } from "node:fs";
import { createElement } from "./oxml/base.js";
import { nsmap } from "./oxml/ns.js";
import { RELATIONSHIP_TYPE as RT } from "./opc/constants.js";
import { XmlPart } from "./opc/package.js";
import { PatternFill } from "./styles/values.js";
import { probeImage } from "./image/probe.js";
import { DrawingPart } from "./parts/drawing.js";
import { getOrAddImagePart, ImagePart } from "./parts/image.js";
import { CT_Sheet } from "./oxml/workbook.js";
import { CT_Cell, CT_Row, CT_SheetData, CT_Worksheet } from "./oxml/worksheet.js";
import type { ChartInfo } from "./oxml/chart.js";
import type { ChartPart } from "./parts/chart.js";
import type { Workbook } from "./workbook.js";
import type { WorksheetPart } from "./parts/worksheet.js";

const MAIN = nsmap.main;

export interface ConditionalRule {
  type: string;
  priority: number;
  operator?: string;
  dxfId?: number;
  formulas: string[];
}

export interface ConditionalFormatting {
  sqref: string;
  rules: ConditionalRule[];
}

export interface DataValidationSpec {
  sqref: string;
  type: string;
  formula1?: string;
  formula2?: string;
  operator?: string;
  allowBlank?: boolean;
  showDropDown?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  promptTitle?: string;
  prompt?: string;
  errorTitle?: string;
  error?: string;
}

export interface HyperlinkSpec {
  ref: string;
  target: string;
  location?: string;
  display?: string;
}

export interface CommentSpec {
  ref: string;
  author: string;
  text: string;
}

export interface CellRange {
  minRow?: number;
  maxRow?: number;
  minCol?: number;
  maxCol?: number;
}

export class Worksheet {
  constructor(
    readonly workbook: Workbook,
    private readonly sheetElm: CT_Sheet,
    readonly part: WorksheetPart,
  ) {}

  /** Sheet tab name. */
  get title(): string {
    return this.sheetElm.name;
  }

  /** "visible" | "hidden" | "veryHidden". */
  get sheetState(): string {
    return this.sheetElm.state;
  }

  private get ct(): CT_Worksheet {
    return this.part.ctWorksheet;
  }

  private get sheetData(): CT_SheetData {
    return this.ct.sheetData;
  }

  /** Internal: the backing `<c>` for `ref`, or null (no mutation). */
  _getCellCt(ref: string): CT_Cell | null {
    const { row } = cellRefToRowCol(ref);
    const rowEl = this.sheetData.findRow(row);
    return rowEl === null ? null : rowEl.findCell(ref);
  }

  /** Internal: find-or-create the backing `<c>` for `ref` (write path). */
  _ensureCellCt(ref: string): CT_Cell {
    const { row } = cellRefToRowCol(ref);
    return this.sheetData.getOrAddRow(row).getOrAddCell(ref);
  }

  /**
   * A cell by A1 coordinate — `cell("B3")` — or by 1-based (row, column) —
   * `cell(3, 2)`. Returns a lazy proxy; nothing is written until you assign.
   */
  cell(coordinate: string): Cell;
  cell(row: number, column: number): Cell;
  cell(a: string | number, b?: number): Cell {
    const ref =
      typeof a === "string" ? normalizeRef(a) : rowColToCellRef(a, b as number);
    return new Cell(this, ref);
  }

  // -- dimensions -------------------------------------------------------

  /** Highest populated row number (0 when the sheet is empty). */
  get maxRow(): number {
    let max = 0;
    for (const row of this.sheetData.rowLst) {
      if (row.r > max) max = row.r;
    }
    return max;
  }

  /** Highest populated column index (0 when the sheet is empty). */
  get maxColumn(): number {
    let max = 0;
    for (const row of this.sheetData.rowLst) {
      for (const c of row.cellLst) {
        const col = cellRefToRowCol(c.ref).col;
        if (col > max) max = col;
      }
    }
    return max;
  }

  /** Bounding A1 range of populated cells, e.g. "A1:C5" ("A1" when empty). */
  get dimensions(): string {
    let minRow = Infinity;
    let minCol = Infinity;
    let maxRow = 0;
    let maxCol = 0;
    for (const row of this.sheetData.rowLst) {
      for (const c of row.cellLst) {
        const { row: r, col } = cellRefToRowCol(c.ref);
        if (r < minRow) minRow = r;
        if (r > maxRow) maxRow = r;
        if (col < minCol) minCol = col;
        if (col > maxCol) maxCol = col;
      }
    }
    if (maxRow === 0) return "A1";
    return `${rowColToCellRef(minRow, minCol)}:${rowColToCellRef(maxRow, maxCol)}`;
  }

  // -- iteration --------------------------------------------------------

  /**
   * Iterate rows as arrays of Cell proxies. Without bounds, spans A1 to the
   * populated max. Empty cells within the range yield null-valued proxies.
   */
  *iterRows(range: CellRange = {}): IterableIterator<Cell[]> {
    const minRow = range.minRow ?? 1;
    const maxRow = range.maxRow ?? this.maxRow;
    const minCol = range.minCol ?? 1;
    const maxCol = range.maxCol ?? this.maxColumn;
    for (let r = minRow; r <= maxRow; r++) {
      const cells: Cell[] = [];
      for (let col = minCol; col <= maxCol; col++) {
        cells.push(new Cell(this, `${getColumnLetter(col)}${r}`));
      }
      yield cells;
    }
  }

  /** Iterate columns as arrays of Cell proxies (transpose of iterRows). */
  *iterCols(range: CellRange = {}): IterableIterator<Cell[]> {
    const minRow = range.minRow ?? 1;
    const maxRow = range.maxRow ?? this.maxRow;
    const minCol = range.minCol ?? 1;
    const maxCol = range.maxCol ?? this.maxColumn;
    for (let col = minCol; col <= maxCol; col++) {
      const cells: Cell[] = [];
      for (let r = minRow; r <= maxRow; r++) {
        cells.push(new Cell(this, `${getColumnLetter(col)}${r}`));
      }
      yield cells;
    }
  }

  /** All rows of the populated area (A1 → max), each an array of Cell proxies. */
  get rows(): Cell[][] {
    return [...this.iterRows()];
  }

  /** All columns of the populated area, each an array of Cell proxies. */
  get columns(): Cell[][] {
    return [...this.iterCols()];
  }

  // -- write ------------------------------------------------------------

  /** Append a row of values at row `maxRow + 1`, filling from column A. */
  append(values: CellValue[]): void {
    const r = this.maxRow + 1;
    for (let i = 0; i < values.length; i++) {
      this.cell(rowColToCellRef(r, i + 1)).value = values[i]!;
    }
  }

  // -- merged cells -----------------------------------------------------

  /** Merged ranges, e.g. ["A5:C5"]. */
  get mergedCells(): string[] {
    return this.ct.mergedRefs;
  }

  mergeCells(range: string): void {
    this.ct.addMerge(range.toUpperCase());
  }

  unmergeCells(range: string): void {
    this.ct.removeMerge(range.toUpperCase());
  }

  // -- column / row dimensions ------------------------------------------

  /** Live accessor for a column's width/hidden state (`ws.column("A")`). */
  column(letter: string): ColumnDimension {
    return new ColumnDimension(this.ct, columnIndexFromString(letter));
  }

  /** Live accessor for a row's height/hidden state (`ws.row(3)`). */
  row(r: number): RowDimension {
    return new RowDimension(this.ct, r);
  }

  // -- structural insert / delete ---------------------------------------

  insertRows(index: number, count = 1): void {
    this.ct.insertRows(index, count);
    this.shiftMerges(index, count, undefined, undefined);
    this.refreshDimension();
  }

  deleteRows(index: number, count = 1): void {
    this.ct.deleteRows(index, count);
    this.shiftMerges(index, -count, undefined, undefined);
    this.refreshDimension();
  }

  insertColumns(index: number, count = 1): void {
    this.ct.insertCols(index, count);
    this.shiftMerges(undefined, undefined, index, count);
    this.refreshDimension();
  }

  deleteColumns(index: number, count = 1): void {
    this.ct.deleteCols(index, count);
    this.shiftMerges(undefined, undefined, index, -count);
    this.refreshDimension();
  }

  // -- freeze panes / autofilter ----------------------------------------

  /** The frozen top-left cell (e.g. "B2"), or null. */
  get freezePanes(): string | null {
    return this.ct.paneTopLeft;
  }
  set freezePanes(ref: string | null) {
    if (ref === null) {
      this.ct.clearFreezePane();
      return;
    }
    const { row, col } = cellRefToRowCol(ref);
    const xSplit = col - 1;
    const ySplit = row - 1;
    if (xSplit === 0 && ySplit === 0) {
      this.ct.clearFreezePane();
      return;
    }
    const activePane = xSplit > 0 && ySplit > 0 ? "bottomRight" : xSplit > 0 ? "topRight" : "bottomLeft";
    this.ct.setFreezePane(xSplit, ySplit, ref.toUpperCase(), activePane);
  }

  get autoFilter(): string | null {
    return this.ct.autoFilterRef;
  }
  set autoFilter(ref: string | null) {
    this.ct.setAutoFilter(ref === null ? null : ref.toUpperCase());
  }

  // -- sheet properties -------------------------------------------------

  set title(name: string) {
    if (name !== this.sheetElm.name && this.workbook.sheetnames.includes(name)) {
      throw new Error(`a worksheet named "${name}" already exists`);
    }
    this.sheetElm.name = name;
  }

  set sheetState(state: "visible" | "hidden" | "veryHidden") {
    this.sheetElm.state = state;
  }

  get tabColor(): string | null {
    return this.ct.tabColorRgb;
  }
  set tabColor(rgb: string | null) {
    this.ct.setTabColorRgb(rgb === null ? null : normalizeRgb(rgb));
  }

  // -- conditional formatting (M5) --------------------------------------

  /** All conditional-formatting blocks with their rules (read). */
  get conditionalFormatting(): ConditionalFormatting[] {
    return this.ct.conditionalFormattingEls.map((cf) => ({
      sqref: cf.getAttr("sqref") ?? "",
      rules: cf.findAll(MAIN, "cfRule").map((rule) => ({
        type: rule.getAttr("type") ?? "",
        priority: Number(rule.getAttr("priority") ?? "0"),
        operator: rule.getAttr("operator") ?? undefined,
        dxfId: rule.getAttr("dxfId") === null ? undefined : Number(rule.getAttr("dxfId")),
        formulas: rule.findAll(MAIN, "formula").map((f) => f.text),
      })),
    }));
  }

  /** Add a `cellIs` conditional-formatting rule with an optional highlight
   * fill. `formula` may be a single value or two (for `between`). */
  addCellIsRule(
    sqref: string,
    opts: { operator: string; formula: string | string[]; fill?: PatternFill },
  ): void {
    const cf = this.ct.addConditionalFormatting(sqref);
    const rule = createElement("main:cfRule", cf);
    cf.appendChild(rule);
    rule.setAttr("type", "cellIs");
    rule.setAttr("priority", String(this.ct.maxCfPriority() + 1));
    rule.setAttr("operator", opts.operator);
    if (opts.fill !== undefined) {
      rule.setAttr("dxfId", String(this.workbook.addDxfFill(opts.fill)));
    }
    for (const f of Array.isArray(opts.formula) ? opts.formula : [opts.formula]) {
      const fEl = createElement("main:formula", rule);
      rule.appendChild(fEl);
      fEl.setText(f);
    }
  }

  // -- data validation (M5) ---------------------------------------------

  get dataValidations(): DataValidationSpec[] {
    return this.ct.dataValidationEls.map((dv) => ({
      sqref: dv.getAttr("sqref") ?? "",
      type: dv.getAttr("type") ?? "",
      formula1: dv.find(MAIN, "formula1")?.text,
      formula2: dv.find(MAIN, "formula2")?.text,
      operator: dv.getAttr("operator") ?? undefined,
      allowBlank: boolAttrOf(dv, "allowBlank"),
      showDropDown: boolAttrOf(dv, "showDropDown"),
      showInputMessage: boolAttrOf(dv, "showInputMessage"),
      showErrorMessage: boolAttrOf(dv, "showErrorMessage"),
    }));
  }

  addDataValidation(spec: DataValidationSpec): void {
    const dv = this.ct.addDataValidation();
    dv.setAttr("sqref", spec.sqref);
    dv.setAttr("type", spec.type);
    if (spec.operator !== undefined) dv.setAttr("operator", spec.operator);
    for (const [k, v] of [
      ["allowBlank", spec.allowBlank],
      ["showDropDown", spec.showDropDown],
      ["showInputMessage", spec.showInputMessage],
      ["showErrorMessage", spec.showErrorMessage],
    ] as const) {
      if (v !== undefined) dv.setAttr(k, v ? "1" : "0");
    }
    for (const [k, v] of [
      ["promptTitle", spec.promptTitle],
      ["prompt", spec.prompt],
      ["errorTitle", spec.errorTitle],
      ["error", spec.error],
    ] as const) {
      if (v !== undefined) dv.setAttr(k, v);
    }
    for (const [tag, val] of [
      ["formula1", spec.formula1],
      ["formula2", spec.formula2],
    ] as const) {
      if (val === undefined) continue;
      const f = createElement(`main:${tag}`, dv);
      dv.appendChild(f);
      f.setText(val);
    }
  }

  // -- hyperlinks (M5) --------------------------------------------------

  get hyperlinks(): HyperlinkSpec[] {
    return this.ct.hyperlinkEls.map((h) => {
      const rId = h.getAttrNS(nsmap.r, "id");
      const location = h.getAttr("location") ?? undefined;
      const target = rId !== null ? this.part.targetRef(rId) : (location ?? "");
      return { ref: h.getAttr("ref") ?? "", target, location, display: h.getAttr("display") ?? undefined };
    });
  }

  /** Add a hyperlink on `ref`. Provide `target` for an external URL (a rel is
   * created) or `location` for an in-workbook target. */
  addHyperlink(
    ref: string,
    opts: { target?: string; location?: string; display?: string },
  ): void {
    const upper = ref.toUpperCase();
    const h = this.ct.addHyperlinkEl(upper);
    if (opts.target !== undefined) {
      const rId = this.part.relateTo(opts.target, RT.HYPERLINK, true);
      h.setAttr("xmlns:r", nsmap.r);
      h.setAttr("r:id", rId);
    }
    if (opts.location !== undefined) h.setAttr("location", opts.location);
    if (opts.display !== undefined) h.setAttr("display", opts.display);
  }

  removeHyperlink(ref: string): void {
    const rId = this.ct.removeHyperlinkEl(ref.toUpperCase());
    if (rId !== null) this.part.dropRel(rId);
  }

  // -- comments (M5, read; authoring deferred) --------------------------

  /** Cell comments attached to this sheet (read; authoring is deferred). */
  get comments(): CommentSpec[] {
    const part = this.commentsPart();
    if (part === null) return [];
    const root = part.root;
    const authors = root.find(MAIN, "authors")?.findAll(MAIN, "author").map((a) => a.text) ?? [];
    const list = root.find(MAIN, "commentList");
    if (list === null) return [];
    return list.findAll(MAIN, "comment").map((c) => {
      const authorId = Number(c.getAttr("authorId") ?? "0");
      const textEl = c.find(MAIN, "text");
      let text = "";
      if (textEl !== null) {
        const directT = textEl.find(MAIN, "t");
        if (directT !== null) text = directT.text;
        else for (const r of textEl.findAll(MAIN, "r")) text += r.find(MAIN, "t")?.text ?? "";
      }
      return { ref: c.getAttr("ref") ?? "", author: authors[authorId] ?? "", text };
    });
  }

  /** Internal: the comments part for this sheet, or null. */
  private commentsPart(): XmlPart | null {
    for (const rel of this.part.rels) {
      if (!rel.isExternal && rel.reltype === RT.COMMENTS) {
        return rel.targetPart as XmlPart;
      }
    }
    return null;
  }

  // -- charts & images (M6) ---------------------------------------------

  private drawingPart(): DrawingPart | null {
    for (const rel of this.part.rels) {
      if (!rel.isExternal && rel.reltype === RT.DRAWING) return rel.targetPart as DrawingPart;
    }
    return null;
  }

  /** Charts anchored on this sheet (read: type, title, series + cached data). */
  get charts(): ChartInfo[] {
    const dp = this.drawingPart();
    if (dp === null) return [];
    const out: ChartInfo[] = [];
    for (const rel of dp.rels) {
      if (!rel.isExternal && rel.reltype === RT.CHART) {
        out.push((rel.targetPart as ChartPart).info);
      }
    }
    return out;
  }

  private imageAnchors(): { ref: string; part: ImagePart }[] {
    const dp = this.drawingPart();
    if (dp === null) return [];
    const out: { ref: string; part: ImagePart }[] = [];
    for (const pic of dp.wsDr.findAllDeep(nsmap.xdr, "pic")) {
      const blip = pic.findAllDeep(nsmap.a, "blip")[0];
      const embed = blip?.getAttrNS(nsmap.r, "embed") ?? null;
      if (embed === null) continue;
      const part = dp.relatedPart(embed);
      if (!(part instanceof ImagePart)) continue;
      const from = pic.parent?.find(nsmap.xdr, "from") ?? null;
      const col = Number(from?.find(nsmap.xdr, "col")?.text ?? "0");
      const row = Number(from?.find(nsmap.xdr, "row")?.text ?? "0");
      out.push({ ref: `${getColumnLetter(col + 1)}${row + 1}`, part });
    }
    return out;
  }

  /** Images embedded on this sheet: anchor cell + media path + content type. */
  get images(): { ref: string; path: string; contentType: string }[] {
    return this.imageAnchors().map((a) => ({
      ref: a.ref,
      path: a.part.partname.uri,
      contentType: a.part.contentType,
    }));
  }

  /** Replace the bytes of the `index`-th embedded image (surgical: only that
   * media entry changes). */
  replaceImage(index: number, source: string | Uint8Array): void {
    const anchors = this.imageAnchors();
    const anchor = anchors[index];
    if (anchor === undefined) throw new Error(`no image at index ${index}`);
    anchor.part.setBlob(toBytes(source));
  }

  /** Embed an image, anchored top-left at `anchorCell`. Creates the sheet's
   * drawing part if it has none. Sizes the picture to the image's pixel dims. */
  addImage(source: string | Uint8Array, anchorCell: string): void {
    const blob = toBytes(source);
    const probe = probeImage(blob);
    const imgPart = getOrAddImagePart(this.part.pkg, blob);
    let dp = this.drawingPart();
    if (dp === null) {
      dp = DrawingPart.createNew(this.part.pkg);
      const rId = this.part.relateTo(dp, RT.DRAWING);
      this.ct.setDrawingRel(rId);
      this.part.pkg.markStructureDirty();
    }
    const embedRId = dp.relateTo(imgPart, RT.IMAGE);
    const { row, col } = cellRefToRowCol(anchorCell.toUpperCase());
    dp.addPicAnchor(embedRId, col - 1, row - 1, probe.pxWidth * 9525, probe.pxHeight * 9525);
  }

  // -- internals --------------------------------------------------------

  private refreshDimension(): void {
    this.ct.dimensionRef = this.dimensions;
  }

  private shiftMerges(
    fromRow: number | undefined,
    dRow: number | undefined,
    fromCol: number | undefined,
    dCol: number | undefined,
  ): void {
    const refs = this.ct.mergedRefs;
    if (refs.length === 0) return;
    const out: string[] = [];
    for (const ref of refs) {
      const b = rangeBoundariesLocal(ref);
      let { minRow, maxRow, minCol, maxCol } = b;
      if (fromRow !== undefined && dRow !== undefined) {
        if (dRow > 0) {
          if (minRow >= fromRow) minRow += dRow;
          if (maxRow >= fromRow) maxRow += dRow;
        } else {
          const removedTo = fromRow - dRow; // fromRow + count
          if (minRow >= fromRow && maxRow < removedTo) continue; // fully deleted
          if (minRow >= removedTo) minRow += dRow;
          else if (minRow >= fromRow) minRow = fromRow;
          if (maxRow >= removedTo) maxRow += dRow;
          else if (maxRow >= fromRow) maxRow = fromRow - 1 < minRow ? minRow : fromRow - 1;
        }
      }
      if (fromCol !== undefined && dCol !== undefined) {
        if (dCol > 0) {
          if (minCol >= fromCol) minCol += dCol;
          if (maxCol >= fromCol) maxCol += dCol;
        } else {
          const removedTo = fromCol - dCol;
          if (minCol >= fromCol && maxCol < removedTo) continue;
          if (minCol >= removedTo) minCol += dCol;
          else if (minCol >= fromCol) minCol = fromCol;
          if (maxCol >= removedTo) maxCol += dCol;
          else if (maxCol >= fromCol) maxCol = fromCol - 1 < minCol ? minCol : fromCol - 1;
        }
      }
      if (minRow > maxRow || minCol > maxCol) continue;
      out.push(
        `${getColumnLetter(minCol)}${minRow}:${getColumnLetter(maxCol)}${maxRow}`,
      );
    }
    this.ct.replaceMerges(out);
  }
}

/** Live view of a `<col>` width/hidden state. */
export class ColumnDimension {
  constructor(
    private readonly ct: CT_Worksheet,
    private readonly idx: number,
  ) {}

  get width(): number | undefined {
    const w = this.ct.findCol(this.idx)?.getAttr("width");
    return w == null ? undefined : Number(w);
  }
  set width(value: number | undefined) {
    const col = this.ct.getOrAddCol(this.idx);
    if (value === undefined) {
      col.removeAttr("width");
      col.removeAttr("customWidth");
    } else {
      col.setAttr("width", String(value));
      col.setAttr("customWidth", "1");
    }
  }

  get hidden(): boolean {
    const h = this.ct.findCol(this.idx)?.getAttr("hidden");
    return h === "1" || h === "true";
  }
  set hidden(value: boolean) {
    const col = this.ct.getOrAddCol(this.idx);
    if (value) col.setAttr("hidden", "1");
    else col.removeAttr("hidden");
  }
}

/** Live view of a `<row>` height/hidden state. */
export class RowDimension {
  constructor(
    private readonly ct: CT_Worksheet,
    private readonly r: number,
  ) {}

  private ensureRow(): CT_Row {
    return this.ct.sheetData.getOrAddRow(this.r);
  }

  get height(): number | undefined {
    const h = this.ct.sheetData.findRow(this.r)?.el.getAttr("ht");
    return h == null ? undefined : Number(h);
  }
  set height(value: number | undefined) {
    const el = this.ensureRow().el;
    if (value === undefined) {
      el.removeAttr("ht");
      el.removeAttr("customHeight");
    } else {
      el.setAttr("ht", String(value));
      el.setAttr("customHeight", "1");
    }
  }

  get hidden(): boolean {
    const h = this.ct.sheetData.findRow(this.r)?.el.getAttr("hidden");
    return h === "1" || h === "true";
  }
  set hidden(value: boolean) {
    const el = this.ensureRow().el;
    if (value) el.setAttr("hidden", "1");
    else el.removeAttr("hidden");
  }
}

function rangeBoundariesLocal(ref: string): {
  minRow: number;
  maxRow: number;
  minCol: number;
  maxCol: number;
} {
  const [a, b] = ref.split(":");
  const p1 = cellRefToRowCol(a!);
  const p2 = b === undefined ? p1 : cellRefToRowCol(b);
  return {
    minRow: Math.min(p1.row, p2.row),
    maxRow: Math.max(p1.row, p2.row),
    minCol: Math.min(p1.col, p2.col),
    maxCol: Math.max(p1.col, p2.col),
  };
}

function normalizeRgb(s: string): string {
  const up = s.toUpperCase();
  if (/^[0-9A-F]{8}$/.test(up)) return up;
  if (/^[0-9A-F]{6}$/.test(up)) return `00${up}`;
  throw new Error(`invalid rgb color: "${s}"`);
}

function boolAttrOf(el: import("./xml/dom.js").XmlElement, name: string): boolean | undefined {
  const v = el.getAttr(name);
  if (v === null) return undefined;
  return v === "1" || v === "true";
}

function toBytes(source: string | Uint8Array): Uint8Array {
  return typeof source === "string" ? new Uint8Array(readFileSync(source)) : source;
}

/** Normalize a user A1 ref: uppercase column, drop absolutes ("$B$3" → "B3"). */
function normalizeRef(ref: string): string {
  const { row, col } = cellRefToRowCol(ref);
  return `${getColumnLetter(col)}${row}`;
}
