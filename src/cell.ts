/**
 * Cell — port of openpyxl/cell/cell.py, with the ts-xlsx fidelity divergence:
 * a Cell is a lazy proxy over (worksheet, coordinate). Reading a cell that has
 * no `<c>` element returns null and creates NOTHING (openpyxl mutates on
 * access; our contract is reads-never-mutate). Writing materializes the element.
 */
import { cellRefToRowCol } from "./util.js";
import { fromExcel, toExcel } from "./datetimes.js";
import { formatValue } from "./numberFormat.js";
import type { CT_Cell } from "./oxml/worksheet.js";
import type { Worksheet } from "./worksheet.js";
import type { Alignment, Border, Fill, Font, Protection } from "./styles/values.js";

/** A cell's readable value. Formulas read back as their "=…" string unless the
 * workbook was opened `{ dataOnly: true }`, in which case the cached value. */
export type CellValue = number | string | boolean | Date | null;

/** openpyxl-style data type marker (plus "f" when the cell holds a formula). */
export type CellType = "n" | "s" | "str" | "b" | "inlineStr" | "e" | "d" | "f";

export class Cell {
  readonly row: number;
  readonly column: number;
  /** A1-style coordinate (column letters uppercased). */
  readonly coordinate: string;

  constructor(
    readonly worksheet: Worksheet,
    coordinate: string,
  ) {
    const { row, col } = cellRefToRowCol(coordinate);
    this.row = row;
    this.column = col;
    this.coordinate = coordinate;
  }

  /** The backing `<c>` element, or null when the cell is empty (no mutation). */
  private get ct(): CT_Cell | null {
    return this.worksheet._getCellCt(this.coordinate);
  }

  /** Column letters of the coordinate ("B" for "B3"). */
  get columnLetter(): string {
    return this.coordinate.replace(/\d+$/, "");
  }

  /** Effective data type: "f" for a formula cell, else the `@t` value. */
  get dataType(): CellType {
    const c = this.ct;
    if (c === null) return "n";
    if (c.fEl !== null) return "f";
    return c.typeAttr as CellType;
  }

  /** True when a formula is stored in this cell. */
  get hasFormula(): boolean {
    return this.ct?.fEl !== null && this.ct?.fEl !== undefined;
  }

  get value(): CellValue {
    const c = this.ct;
    if (c === null) return null;

    if (c.fEl !== null && !this.worksheet.workbook.dataOnly) {
      return `=${c.formula ?? ""}`;
    }

    switch (c.typeAttr) {
      case "s": {
        const raw = c.vText;
        if (raw === null || raw === "") return null;
        return this.worksheet.workbook.sharedStringText(Number(raw));
      }
      case "inlineStr":
        return c.inlineText ?? "";
      case "str":
        return c.vText ?? "";
      case "b":
        return c.vText === "1";
      case "e":
        return c.vText ?? "";
      case "d":
        return c.vText ? new Date(c.vText) : null;
      default: {
        // "n" (number) — possibly a serial date governed by its number format
        const raw = c.vText;
        if (raw === null || raw === "") return null;
        const num = Number(raw);
        if (this.worksheet.workbook.isDateStyle(c.styleIndex)) {
          return fromExcel(num, this.worksheet.workbook.date1904);
        }
        return num;
      }
    }
  }

  /**
   * Assign a value. Dispatch by JS type: number → numeric cell; boolean →
   * `t="b"`; Date → serial + a date number format; string starting with "=" →
   * a stored formula (sets fullCalcOnLoad, drops the stale calcChain); any other
   * string → an interned shared string; null/undefined → clears the cell.
   */
  set value(v: CellValue) {
    const wb = this.worksheet.workbook;

    if (v === null || v === undefined) {
      const existing = this.worksheet._getCellCt(this.coordinate);
      if (existing !== null) {
        existing.clearValue();
        existing.typeAttr = "n";
      }
      return;
    }

    const c = this.worksheet._ensureCellCt(this.coordinate);

    if (typeof v === "number") {
      c.clearValue();
      c.typeAttr = "n";
      c.getOrAddV().setText(numberToXml(v));
    } else if (typeof v === "boolean") {
      c.clearValue();
      c.typeAttr = "b";
      c.getOrAddV().setText(v ? "1" : "0");
    } else if (v instanceof Date) {
      const serial = toExcel(v, wb.date1904);
      c.clearValue();
      c.typeAttr = "n";
      c.getOrAddV().setText(numberToXml(serial));
      const hasTime =
        v.getUTCHours() !== 0 ||
        v.getUTCMinutes() !== 0 ||
        v.getUTCSeconds() !== 0 ||
        v.getUTCMilliseconds() !== 0;
      c.styleIndex = wb.applyDateFormat(c.styleIndex, hasTime);
    } else if (typeof v === "string") {
      if (v.startsWith("=")) {
        c.clearValue();
        c.typeAttr = "n";
        c.setFormula(v.slice(1));
        wb.setFullCalcOnLoad();
        wb.invalidateCalcChain();
      } else {
        const idx = wb.internString(v);
        c.clearValue();
        c.typeAttr = "s";
        c.getOrAddV().setText(String(idx));
      }
    }
  }

  // -- styles (M3) — read is non-mutating; set finds-or-extends an xf ----

  /** Cell style index (`@s`) into the cellXfs chain. */
  get styleIndex(): number {
    return this.ct?.styleIndex ?? 0;
  }

  private applyStyle(next: (styleIndex: number) => number): void {
    const c = this.worksheet._ensureCellCt(this.coordinate);
    c.styleIndex = next(c.styleIndex);
  }

  /** Number-format code applied to this cell (e.g. "#,##0.00"). */
  get numberFormat(): string {
    return this.worksheet.workbook.formatCodeForStyle(this.styleIndex);
  }
  set numberFormat(code: string) {
    this.applyStyle((s) => this.worksheet.workbook.applyNumberFormat(s, code));
  }

  get font(): Font {
    return this.worksheet.workbook.fontForStyle(this.styleIndex);
  }
  set font(font: Font) {
    this.applyStyle((s) => this.worksheet.workbook.applyFont(s, font));
  }

  get fill(): Fill {
    return this.worksheet.workbook.fillForStyle(this.styleIndex);
  }
  set fill(fill: Fill) {
    this.applyStyle((s) => this.worksheet.workbook.applyFill(s, fill));
  }

  get border(): Border {
    return this.worksheet.workbook.borderForStyle(this.styleIndex);
  }
  set border(border: Border) {
    this.applyStyle((s) => this.worksheet.workbook.applyBorder(s, border));
  }

  get alignment(): Alignment {
    return this.worksheet.workbook.alignmentForStyle(this.styleIndex);
  }
  set alignment(alignment: Alignment) {
    this.applyStyle((s) => this.worksheet.workbook.applyAlignment(s, alignment));
  }

  get protection(): Protection {
    return this.worksheet.workbook.protectionForStyle(this.styleIndex);
  }
  set protection(protection: Protection) {
    this.applyStyle((s) => this.worksheet.workbook.applyProtection(s, protection));
  }

  /** Name of the applied named cell style ("Normal", "Title", …). Assigning
   * applies an existing named style from the workbook. */
  get styleName(): string {
    return this.worksheet.workbook.styleNameForStyle(this.styleIndex);
  }
  set styleName(name: string) {
    this.applyStyle(() => this.worksheet.workbook.applyNamedStyle(name));
  }

  /** The external target / location of a hyperlink on this cell, or null.
   * Assigning a string adds an external-URL hyperlink. */
  get hyperlink(): string | null {
    const h = this.worksheet.hyperlinks.find((x) => x.ref === this.coordinate);
    return h === undefined ? null : h.target;
  }
  set hyperlink(target: string | null) {
    if (target === null) {
      this.worksheet.removeHyperlink(this.coordinate);
    } else {
      this.worksheet.addHyperlink(this.coordinate, { target });
    }
  }

  /** The comment attached to this cell (read), or null. */
  get comment(): { author: string; text: string } | null {
    const c = this.worksheet.comments.find((x) => x.ref === this.coordinate);
    return c === undefined ? null : { author: c.author, text: c.text };
  }

  /** The value rendered through this cell's number format, as Excel displays
   * it (formula cells with no cached value render empty in dataOnly mode). */
  get displayValue(): string {
    const v = this.value;
    if (typeof v === "string" && v.startsWith("=")) return v; // unresolved formula text
    return formatValue(v, this.numberFormat, this.worksheet.workbook.date1904);
  }
}

/** Serialize a number the way Excel stores it in `<v>` (compact decimal). */
function numberToXml(n: number): string {
  if (!Number.isFinite(n)) throw new Error(`cannot store non-finite number ${n}`);
  return String(n);
}
