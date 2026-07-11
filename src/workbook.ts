/**
 * Workbook — merged port of openpyxl/workbook/workbook.py + the load_workbook
 * factory. `Workbook.open()` replaces the factory (a source is always required).
 * `save`/`toBuffer` are async (they touch the zip); the object model is sync.
 */
import { XlsxError, KeyLookupError } from "./exc.js";
import { OpcPackage, Part } from "./opc/package.js";
import { RELATIONSHIP_TYPE as RT } from "./opc/constants.js";
import { nsmap } from "./oxml/ns.js";
import { CT_Workbook } from "./oxml/workbook.js";
import { FORMAT_GENERAL, isDateFormat } from "./numberFormats.js";
import { Alignment, Border, Fill, Font, PatternFill, Protection } from "./styles/values.js";
import { Worksheet } from "./worksheet.js";
import { SharedStringsPart } from "./parts/sharedstrings.js";
import { WorksheetPart } from "./parts/worksheet.js";
import type { WorkbookPart } from "./parts/workbook.js";
import type { StylesPart } from "./parts/styles.js";
import type { CorePropertiesPart, CoreProperties } from "./parts/coreprops.js";
import type { CT_Sst } from "./oxml/sharedstrings.js";
import type { CT_Stylesheet } from "./oxml/styles.js";

const MAIN = nsmap.main;

export interface WorkbookOpenOptions {
  /** When true, formula cells read back their cached value, not the "=…" text. */
  dataOnly?: boolean;
}

export class Workbook {
  /** Set by open(); when true, formula cells return their cached `<v>`. */
  dataOnly = false;

  #worksheets: Worksheet[] | undefined;
  #sst: CT_Sst | null | undefined;
  #stylesheet: CT_Stylesheet | null | undefined;

  constructor(readonly part: WorkbookPart) {}

  /** Open an existing .xlsx from a path, Uint8Array, or ArrayBuffer. */
  static async open(
    source: string | Uint8Array | ArrayBuffer,
    opts: WorkbookOpenOptions = {},
  ): Promise<Workbook> {
    const pkg = await OpcPackage.open(source);
    const main = pkg.mainDocumentPart;
    // WorkbookPart (registered in src/index.ts) exposes `.workbook`.
    const wb = (main as { workbook?: Workbook }).workbook;
    if (!(wb instanceof Workbook)) {
      throw new XlsxError(
        `main document part is not a workbook (content type: ${main.contentType})`,
      );
    }
    wb.dataOnly = opts.dataOnly ?? false;
    return wb;
  }

  private get ctWorkbook(): CT_Workbook {
    return new CT_Workbook(this.part.root);
  }

  /** Worksheets in tab order. */
  get worksheets(): Worksheet[] {
    if (this.#worksheets === undefined) {
      this.#worksheets = this.ctWorkbook.sheetLst.map((sheet) => {
        const rId = sheet.rId;
        if (rId === null) throw new XlsxError(`sheet "${sheet.name}" has no r:id`);
        const part = this.part.relatedPart(rId) as WorksheetPart;
        return new Worksheet(this, sheet, part);
      });
    }
    return this.#worksheets;
  }

  /** Sheet names in tab order. */
  get sheetnames(): string[] {
    return this.worksheets.map((ws) => ws.title);
  }

  /** The active worksheet (workbookView/@activeTab, defaulting to the first). */
  get active(): Worksheet {
    const view = this.part.root.find(MAIN, "bookViews")?.find(MAIN, "workbookView");
    const raw = view?.getAttr("activeTab");
    const idx = raw ? Number(raw) : 0;
    return this.worksheets[idx] ?? this.worksheets[0]!;
  }

  /** Worksheet by name; throws KeyLookupError if absent. */
  get(name: string): Worksheet {
    const ws = this.worksheets.find((w) => w.title === name);
    if (ws === undefined) throw new KeyLookupError(`no worksheet named "${name}"`);
    return ws;
  }

  /** Whether the workbook uses the 1904 date system. */
  get date1904(): boolean {
    return this.ctWorkbook.date1904;
  }

  // -- sheet management (M4) --------------------------------------------

  /** Create a new empty worksheet, optionally at `index`. */
  createSheet(name?: string, index?: number): Worksheet {
    const finalName = name ?? this.uniqueSheetName("Sheet");
    if (this.sheetnames.includes(finalName)) {
      throw new XlsxError(`a worksheet named "${finalName}" already exists`);
    }
    const part = WorksheetPart.createNew(this.part.pkg);
    const rId = this.part.relateTo(part, RT.WORKSHEET);
    this.ctWorkbook.addSheet(finalName, this.ctWorkbook.nextSheetId(), rId, index);
    this.part.pkg.markStructureDirty();
    this.#worksheets = undefined;
    return this.get(finalName);
  }

  /** Remove a worksheet by name or reference. Throws when removing the last one. */
  removeSheet(target: string | Worksheet): void {
    const name = typeof target === "string" ? target : target.title;
    if (this.worksheets.length <= 1) throw new XlsxError("cannot remove the only worksheet");
    const rId = this.ctWorkbook.removeSheet(name);
    if (rId === null) throw new KeyLookupError(`no worksheet named "${name}"`);
    this.part.dropRel(rId);
    this.part.pkg.markStructureDirty();
    this.#worksheets = undefined;
  }

  /** Copy a worksheet's contents into a new sheet (drawing/hyperlink rels on
   * the copy are not carried in v1 — see the deferred list). */
  copyWorksheet(source: Worksheet, newName?: string): Worksheet {
    const finalName = newName ?? this.uniqueSheetName(`${source.title} Copy`);
    const part = WorksheetPart.createCopy(this.part.pkg, source.part.blob);
    const rId = this.part.relateTo(part, RT.WORKSHEET);
    this.ctWorkbook.addSheet(finalName, this.ctWorkbook.nextSheetId(), rId);
    this.part.pkg.markStructureDirty();
    this.#worksheets = undefined;
    return this.get(finalName);
  }

  /** Move a worksheet to absolute position `index` in the tab order. */
  moveSheet(target: string | Worksheet, index: number): void {
    const name = typeof target === "string" ? target : target.title;
    this.ctWorkbook.moveSheet(name, index);
    this.#worksheets = undefined;
  }

  private uniqueSheetName(base: string): string {
    const taken = new Set(this.sheetnames);
    if (!taken.has(base)) return base;
    for (let n = 1; ; n++) {
      const candidate = `${base}${n}`;
      if (!taken.has(candidate)) return candidate;
    }
  }

  // -- shared strings & styles (used by Cell reads) ---------------------

  private relatedPartOrNull(reltype: string): Part | null {
    for (const rel of this.part.rels) {
      if (!rel.isExternal && rel.reltype === reltype) return rel.targetPart;
    }
    return null;
  }

  private get sst(): CT_Sst | null {
    if (this.#sst === undefined) {
      const part = this.relatedPartOrNull(RT.SHARED_STRINGS) as SharedStringsPart | null;
      this.#sst = part === null ? null : part.sst;
    }
    return this.#sst;
  }

  private get stylesheet(): CT_Stylesheet | null {
    if (this.#stylesheet === undefined) {
      const part = this.relatedPartOrNull(RT.STYLES) as StylesPart | null;
      this.#stylesheet = part === null ? null : part.stylesheet;
    }
    return this.#stylesheet;
  }

  /** Text of shared string `index`; throws if the table is missing/short. */
  sharedStringText(index: number): string {
    const sst = this.sst;
    if (sst === null) {
      throw new XlsxError(
        `cell references shared string ${index} but the workbook has no sharedStrings part`,
      );
    }
    const text = sst.textAt(index);
    if (text === undefined) throw new XlsxError(`shared string index ${index} out of range`);
    return text;
  }

  /** Number-format code applied by cell style index `s` (General if no styles). */
  formatCodeForStyle(s: number): string {
    const ss = this.stylesheet;
    return ss === null ? FORMAT_GENERAL : ss.formatCodeForStyle(s);
  }

  /** Whether style index `s` applies a date/time number format. */
  isDateStyle(s: number): boolean {
    return isDateFormat(this.formatCodeForStyle(s));
  }

  // -- write helpers (M2) -----------------------------------------------

  #stringIndex: Map<string, number> | undefined;

  /** Find-or-add `text` in the shared string table, returning its index.
   * Creates the sharedStrings part if the workbook lacks one. */
  internString(text: string): number {
    let part = this.relatedPartOrNull(RT.SHARED_STRINGS) as SharedStringsPart | null;
    if (part === null) {
      part = SharedStringsPart.createNew(this.part.pkg);
      this.part.relateTo(part, RT.SHARED_STRINGS);
      this.#sst = undefined; // invalidate the read cache
      this.#stringIndex = undefined;
    }
    const sst = part.sst;
    if (this.#stringIndex === undefined) {
      this.#stringIndex = new Map();
      const items = sst.siLst;
      for (let i = 0; i < items.length; i++) this.#stringIndex.set(items[i]!.text, i);
    }
    const existing = this.#stringIndex.get(text);
    if (existing !== undefined) return existing;
    const idx = sst.addPlain(text);
    this.#stringIndex.set(text, idx);
    return idx;
  }

  /** Stylesheet, required; throws when a workbook genuinely has no styles part. */
  private ensureStylesheet(): CT_Stylesheet {
    const part = this.relatedPartOrNull(RT.STYLES) as StylesPart | null;
    if (part === null) {
      throw new XlsxError("workbook has no styles part; cannot apply a number format");
    }
    this.#stylesheet = part.stylesheet;
    return this.#stylesheet;
  }

  /** Apply a default date/datetime number format to `styleIndex`, returning the
   * (reused or new) cell-xf index. No-op-returns the index unchanged when the
   * current format is already a date format (respect an explicit user format). */
  applyDateFormat(styleIndex: number, hasTime: boolean): number {
    if (isDateFormat(this.formatCodeForStyle(styleIndex))) return styleIndex;
    const ss = this.ensureStylesheet();
    const code = hasTime ? "yyyy-mm-dd h:mm:ss" : "yyyy-mm-dd";
    const numFmtId = ss.getOrAddNumFmt(code);
    return ss.getOrAddXfWithNumFmt(styleIndex, numFmtId);
  }

  // -- style reads (M3) -------------------------------------------------

  fontForStyle(s: number): Font {
    const ss = this.stylesheet;
    return ss === null ? new Font() : ss.fontForStyle(s);
  }
  fillForStyle(s: number): Fill {
    const ss = this.stylesheet;
    return ss === null ? new PatternFill() : ss.fillForStyle(s);
  }
  borderForStyle(s: number): Border {
    const ss = this.stylesheet;
    return ss === null ? new Border() : ss.borderForStyle(s);
  }
  alignmentForStyle(s: number): Alignment {
    const ss = this.stylesheet;
    return ss === null ? new Alignment() : ss.alignmentForStyle(s);
  }
  protectionForStyle(s: number): Protection {
    const ss = this.stylesheet;
    return ss === null ? new Protection() : ss.protectionForStyle(s);
  }

  // -- style writes (M3): find-or-add table entry, then find-or-add xf ---

  applyFont(s: number, font: Font): number {
    const ss = this.ensureStylesheet();
    return ss.getOrAddXf(s, { fontId: ss.getOrAddFont(font) });
  }
  applyFill(s: number, fill: Fill): number {
    const ss = this.ensureStylesheet();
    return ss.getOrAddXf(s, { fillId: ss.getOrAddFill(fill) });
  }
  applyBorder(s: number, border: Border): number {
    const ss = this.ensureStylesheet();
    return ss.getOrAddXf(s, { borderId: ss.getOrAddBorder(border) });
  }
  applyAlignment(s: number, alignment: Alignment): number {
    return this.ensureStylesheet().getOrAddXf(s, { alignment });
  }
  applyProtection(s: number, protection: Protection): number {
    return this.ensureStylesheet().getOrAddXf(s, { protection });
  }
  applyNumberFormat(s: number, code: string): number {
    const ss = this.ensureStylesheet();
    return ss.getOrAddXf(s, { numFmtId: ss.getOrAddNumFmt(code) });
  }

  /** Names of the workbook's named cell styles. */
  get namedStyleNames(): string[] {
    const ss = this.stylesheet;
    return ss === null ? [] : ss.namedStyleNames;
  }
  styleNameForStyle(s: number): string {
    const ss = this.stylesheet;
    return ss === null ? "Normal" : ss.styleNameForStyle(s);
  }
  applyNamedStyle(name: string): number {
    return this.ensureStylesheet().applyNamedStyle(name);
  }

  /** Add a differential-format fill for conditional formatting; returns dxfId. */
  addDxfFill(fill: PatternFill): number {
    return this.ensureStylesheet().getOrAddDxfFill(fill);
  }

  // -- defined names (M5) -----------------------------------------------

  /** Workbook defined names: name → refers-to string. */
  get definedNames(): Map<string, string> {
    return this.ctWorkbook.definedNamesMap;
  }

  setDefinedName(name: string, refersTo: string): void {
    this.ctWorkbook.setDefinedName(name, refersTo);
  }

  removeDefinedName(name: string): void {
    this.ctWorkbook.removeDefinedName(name);
  }

  /** Force a full recalculation when Excel next opens the workbook. */
  setFullCalcOnLoad(): void {
    this.ctWorkbook.setFullCalcOnLoad(true);
  }

  /** Drop the (now-stale) calcChain part so Excel rebuilds it after a formula
   * edit — the openpyxl policy, made surgical: only the rel + part are removed. */
  invalidateCalcChain(): void {
    for (const rel of [...this.part.rels]) {
      if (!rel.isExternal && rel.reltype === RT.CALC_CHAIN) {
        this.part.dropRel(rel.rId);
        this.part.pkg.markStructureDirty();
        return;
      }
    }
  }

  // -- save -------------------------------------------------------------

  /** Core document properties (docProps/core.xml) — read/write. */
  get coreProperties(): CoreProperties {
    const part = this.relatedPartFromPackage(RT.CORE_PROPERTIES);
    if (part === null) throw new XlsxError("workbook has no core properties part");
    return (part as CorePropertiesPart).coreProperties;
  }

  private relatedPartFromPackage(reltype: string): Part | null {
    for (const rel of this.part.pkg.rels) {
      if (!rel.isExternal && rel.reltype === reltype) return rel.targetPart;
    }
    return null;
  }

  async save(path: string): Promise<void> {
    await this.part.pkg.save(path);
  }

  async toBuffer(): Promise<Uint8Array> {
    return this.part.pkg.toBuffer();
  }
}
