/**
 * Units and coordinate helpers — port of openpyxl/utils/units.py and
 * openpyxl/utils/cell.py.
 *
 * SpreadsheetML exposes lengths in mixed units: column widths in *character*
 * units, row heights in *points*, and drawing anchors in EMU. openpyxl surfaces
 * width/height as plain floats (no wrapper), so ts-xlsx does too — the EMU
 * family below exists only for the drawing/image layer (M6).
 */

// -- EMU length family (drawings/images only) ---------------------------------
export type Emu = number & { readonly __emu?: unique symbol };

/** 1 inch = 914400 EMU; 1 pt = 12700 EMU; 1 cm = 360000 EMU. */
export const Emu = (n: number): Emu => Math.trunc(n) as Emu;
export const Inches = (n: number): Emu => Emu(n * 914400);
export const Pt = (n: number): Emu => Emu(n * 12700);
export const Cm = (n: number): Emu => Emu(n * 360000);
export const Mm = (n: number): Emu => Emu(n * 36000);
/** Pixels at 96 DPI (Excel's screen unit). 1 px = 9525 EMU. */
export const Pixels = (n: number): Emu => Emu(n * 9525);

export const Length = {
  inches: (v: Emu): number => v / 914400,
  pt: (v: Emu): number => v / 12700,
  cm: (v: Emu): number => v / 360000,
  mm: (v: Emu): number => v / 36000,
  emu: (v: Emu): number => v,
  pixels: (v: Emu): number => Math.round(v / 9525),
};

// -- column-letter <-> index (1-based, bijective base-26) ---------------------

/** Largest column Excel supports: XFD = 16384. */
export const MAX_COLUMN = 16384;
/** Largest row Excel supports: 1048576. */
export const MAX_ROW = 1048576;

const COLUMN_LETTER_RE = /^[A-Za-z]{1,3}$/;

/** "A" → 1, "Z" → 26, "AA" → 27, "XFD" → 16384. */
export function columnIndexFromString(letters: string): number {
  if (!COLUMN_LETTER_RE.test(letters)) {
    throw new Error(`invalid column letters: "${letters}"`);
  }
  let n = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    n = n * 26 + (upper.charCodeAt(i) - 64);
  }
  if (n < 1 || n > MAX_COLUMN) throw new Error(`column out of range: "${letters}"`);
  return n;
}

/** 1 → "A", 26 → "Z", 27 → "AA", 16384 → "XFD". */
export function getColumnLetter(index: number): string {
  if (!Number.isInteger(index) || index < 1 || index > MAX_COLUMN) {
    throw new Error(`column index out of range: ${index}`);
  }
  let letters = "";
  let n = index;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.trunc((n - 1) / 26);
  }
  return letters;
}

// -- A1 <-> (row, col) --------------------------------------------------------

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$/;

/** "$B$3" / "B3" → { column: "B", row: 3 } (absolute markers stripped). */
export function coordinateFromString(coord: string): { column: string; row: number } {
  const m = CELL_RE.exec(coord);
  if (m === null) throw new Error(`invalid cell coordinate: "${coord}"`);
  const row = Number(m[2]);
  if (row > MAX_ROW) throw new Error(`row out of range: "${coord}"`);
  return { column: m[1]!.toUpperCase(), row };
}

/** "B3" → { row: 3, col: 2 } (both 1-based). */
export function cellRefToRowCol(coord: string): { row: number; col: number } {
  const { column, row } = coordinateFromString(coord);
  return { row, col: columnIndexFromString(column) };
}

/** (row=3, col=2) → "B3". */
export function rowColToCellRef(row: number, col: number): string {
  if (!Number.isInteger(row) || row < 1 || row > MAX_ROW) {
    throw new Error(`row out of range: ${row}`);
  }
  return `${getColumnLetter(col)}${row}`;
}

/** "$B$3" → "$B$3" absolute form; ("B", 3) style also accepted via strings. */
export function absoluteCoordinate(coord: string): string {
  const { column, row } = coordinateFromString(coord);
  return `$${column}$${row}`;
}

// -- ranges -------------------------------------------------------------------

export interface RangeBoundaries {
  minCol: number;
  minRow: number;
  maxCol: number;
  maxRow: number;
}

/**
 * "A1:C5" → { minCol:1, minRow:1, maxCol:3, maxRow:5 }. A single cell "B2"
 * yields equal min/max. Column-only ("A:C") and row-only ("1:5") ranges are
 * not supported here (they need sheet dimensions to resolve) and throw.
 */
export function rangeBoundaries(range: string): RangeBoundaries {
  const parts = range.split(":");
  if (parts.length === 1) {
    const { row, col } = cellRefToRowCol(parts[0]!);
    return { minCol: col, minRow: row, maxCol: col, maxRow: row };
  }
  if (parts.length !== 2) throw new Error(`invalid range: "${range}"`);
  const a = cellRefToRowCol(parts[0]!);
  const b = cellRefToRowCol(parts[1]!);
  return {
    minCol: Math.min(a.col, b.col),
    minRow: Math.min(a.row, b.row),
    maxCol: Math.max(a.col, b.col),
    maxRow: Math.max(a.row, b.row),
  };
}

/** { minCol:1, minRow:1, maxCol:3, maxRow:5 } → "A1:C5". */
export function boundariesToRange(b: RangeBoundaries): string {
  const tl = rowColToCellRef(b.minRow, b.minCol);
  const br = rowColToCellRef(b.maxRow, b.maxCol);
  return tl === br ? tl : `${tl}:${br}`;
}
