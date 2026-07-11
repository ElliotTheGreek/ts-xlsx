import { describe, expect, it } from "vitest";
import { Workbook } from "../src/index.js";
import { bytesEqual, fixturePath, zipEntries, zipEntriesOf } from "./helpers/zip.js";

describe("M1 workbook / worksheet navigation", () => {
  it("lists sheets, active, get(name), and sheet state", async () => {
    const wb = await Workbook.open(fixturePath("basic.xlsx"));
    expect(wb.sheetnames).toEqual(["Data", "Notes", "Hidden"]);
    expect(wb.active.title).toBe("Data");
    expect(wb.get("Notes").title).toBe("Notes");
    expect(wb.get("Hidden").sheetState).toBe("hidden");
    expect(() => wb.get("Nope")).toThrow();
  });
});

describe("M1 cell reads — inline strings / numbers / bools / merges", () => {
  it("reads typed values from the openpyxl basic workbook", async () => {
    const wb = await Workbook.open(fixturePath("basic.xlsx"));
    const ws = wb.get("Data");
    expect(ws.cell("A1").value).toBe("Name");
    expect(ws.cell("B2").value).toBe(42);
    expect(ws.cell(2, 2).value).toBe(42); // (row, col) form
    expect(ws.cell("B3").value).toBe(17.5);
    expect(ws.cell("C1").value).toBe(true);
    expect(ws.cell("A5").value).toBe("merged title");
    expect(ws.cell("Z99").value).toBeNull(); // empty cell, no element
    expect(ws.dimensions).toBe("A1:C5");
    expect(ws.maxRow).toBe(5);
    expect(ws.maxColumn).toBe(3);
    expect(ws.cell("B2").dataType).toBe("n");
    expect(ws.cell("C1").dataType).toBe("b");
  });
});

describe("M1 cell reads — dates, unicode, shared strings", () => {
  it("reads dates via number format and unicode text (types.xlsx)", async () => {
    const wb = await Workbook.open(fixturePath("types.xlsx"));
    const ws = wb.active;
    const dt = ws.cell("A1").value as Date;
    expect(dt).toBeInstanceOf(Date);
    expect([dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate()]).toEqual([2026, 6, 11]);
    expect(dt.getUTCHours()).toBe(9);
    expect(dt.getUTCMinutes()).toBe(30);
    const d2 = ws.cell("A2").value as Date;
    expect([d2.getUTCFullYear(), d2.getUTCMonth(), d2.getUTCDate()]).toEqual([2026, 0, 1]);
    expect(ws.cell("A3").value).toBe(true);
    expect(ws.cell("A4").value).toBe(false);
    expect(ws.cell("A8").value).toBe(3.14159);
    expect(ws.cell("A9").value).toBe("unicode: café ☕ 数");
  });

  it("dereferences the shared string table (t='s') and cached formulas", async () => {
    const wb = await Workbook.open(fixturePath("shared_strings.xlsx"));
    const ws = wb.active;
    expect(ws.cell("A1").value).toBe("Hello");
    expect(ws.cell("B1").value).toBe("World");
    expect(ws.cell("A2").value).toBe(" spaced ");
    expect(ws.cell("B2").value).toBe(42);
    // formula cell: default reads the "=…" text; dataType is "f"
    expect(ws.cell("A3").value).toBe("=SUM(B2:B2)");
    expect(ws.cell("A3").dataType).toBe("f");
    // date-formatted numeric via style index 1 (numFmtId 14)
    const d = ws.cell("B3").value as Date;
    expect([d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()]).toEqual([2026, 0, 1]);
  });

  it("dataOnly mode returns the cached formula value", async () => {
    const wb = await Workbook.open(fixturePath("shared_strings.xlsx"), { dataOnly: true });
    expect(wb.active.cell("A3").value).toBe(42);
  });

  it("reads a stored formula string from the openpyxl workbook", async () => {
    const wb = await Workbook.open(fixturePath("formulas.xlsx"));
    expect(wb.active.cell("B1").value).toBe("=SUM(A1:A5)");
  });
});

describe("M1 fidelity — reads never mutate", () => {
  it("open → read many cells → save is byte-identical to the source", async () => {
    const path = fixturePath("basic.xlsx");
    const original = await zipEntriesOf(path);
    const wb = await Workbook.open(path);
    // touch a broad swath of the read surface
    for (const ws of wb.worksheets) {
      void ws.dimensions;
      void ws.maxRow;
      void ws.maxColumn;
      for (const row of ws.iterRows()) for (const c of row) void c.value;
    }
    const saved = await zipEntries(await wb.toBuffer());
    expect([...saved.keys()].sort()).toEqual([...original.keys()].sort());
    for (const [name, bytes] of original) {
      expect(bytesEqual(saved.get(name)!, bytes), `${name} changed on read-only round-trip`).toBe(
        true,
      );
    }
  });
});
