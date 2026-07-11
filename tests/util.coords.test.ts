import { describe, expect, it } from "vitest";
import {
  columnIndexFromString,
  getColumnLetter,
  cellRefToRowCol,
  rowColToCellRef,
  coordinateFromString,
  rangeBoundaries,
  boundariesToRange,
  MAX_COLUMN,
} from "../src/util.js";
import { fromExcel, toExcel } from "../src/datetimes.js";
import { isDateFormat } from "../src/numberFormats.js";

describe("column letter <-> index", () => {
  it("round-trips the boundary values", () => {
    for (const [letters, idx] of [
      ["A", 1],
      ["Z", 26],
      ["AA", 27],
      ["AZ", 52],
      ["BA", 53],
      ["ZZ", 702],
      ["AAA", 703],
      ["XFD", MAX_COLUMN],
    ] as const) {
      expect(columnIndexFromString(letters)).toBe(idx);
      expect(getColumnLetter(idx)).toBe(letters);
    }
  });

  it("is case-insensitive on input", () => {
    expect(columnIndexFromString("aa")).toBe(27);
  });

  it("rejects out-of-range and malformed input", () => {
    expect(() => getColumnLetter(0)).toThrow();
    expect(() => getColumnLetter(MAX_COLUMN + 1)).toThrow();
    expect(() => columnIndexFromString("1A")).toThrow();
  });
});

describe("A1 <-> (row, col)", () => {
  it("parses and builds coordinates", () => {
    expect(cellRefToRowCol("B3")).toEqual({ row: 3, col: 2 });
    expect(cellRefToRowCol("$B$3")).toEqual({ row: 3, col: 2 });
    expect(rowColToCellRef(3, 2)).toBe("B3");
    expect(coordinateFromString("AA100")).toEqual({ column: "AA", row: 100 });
  });

  it("computes range boundaries both ways", () => {
    expect(rangeBoundaries("A1:C5")).toEqual({ minCol: 1, minRow: 1, maxCol: 3, maxRow: 5 });
    expect(rangeBoundaries("C5:A1")).toEqual({ minCol: 1, minRow: 1, maxCol: 3, maxRow: 5 });
    expect(rangeBoundaries("B2")).toEqual({ minCol: 2, minRow: 2, maxCol: 2, maxRow: 2 });
    expect(boundariesToRange({ minCol: 1, minRow: 1, maxCol: 3, maxRow: 5 })).toBe("A1:C5");
    expect(boundariesToRange({ minCol: 2, minRow: 2, maxCol: 2, maxRow: 2 })).toBe("B2");
  });
});

describe("Excel serial dates", () => {
  it("round-trips a datetime through the 1900 system", () => {
    const serial = 46214.39583333334; // 2026-07-11 09:30
    const d = fromExcel(serial);
    expect(d.getUTCFullYear()).toBe(2026);
    expect(d.getUTCMonth()).toBe(6);
    expect(d.getUTCDate()).toBe(11);
    expect(d.getUTCHours()).toBe(9);
    expect(d.getUTCMinutes()).toBe(30);
    expect(toExcel(d)).toBeCloseTo(serial, 6);
  });

  it("handles a whole-day serial", () => {
    const d = fromExcel(46023);
    expect([d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()]).toEqual([2026, 0, 1]);
    expect(toExcel(d)).toBe(46023);
  });

  it("classifies date vs non-date formats (openpyxl parity)", () => {
    expect(isDateFormat("mm-dd-yy")).toBe(true);
    expect(isDateFormat("yyyy-mm-dd h:mm:ss")).toBe(true);
    expect(isDateFormat("General")).toBe(false);
    expect(isDateFormat("#,##0.00")).toBe(false);
    // a quoted literal containing 'd' must not be treated as a date
    expect(isDateFormat('"day"0')).toBe(false);
  });
});
