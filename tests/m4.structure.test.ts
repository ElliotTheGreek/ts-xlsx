import { describe, expect, it } from "vitest";
import { Workbook } from "../src/index.js";
import { fixturePath } from "./helpers/zip.js";

async function reopen(wb: Workbook): Promise<Workbook> {
  return Workbook.open(await wb.toBuffer());
}

describe("M4 merged cells", () => {
  it("reads, adds, and removes merges", async () => {
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    let ws = wb.get("Data");
    expect(ws.mergedCells).toEqual(["A5:C5"]);
    ws.mergeCells("A7:B8");
    ws.unmergeCells("A5:C5");
    wb = await reopen(wb);
    ws = wb.get("Data");
    expect(ws.mergedCells).toEqual(["A7:B8"]);
  });
});

describe("M4 column / row dimensions", () => {
  it("reads and writes width, height, and hidden", async () => {
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    let ws = wb.get("Data");
    expect(ws.column("A").width).toBe(20); // set in the fixture
    ws.column("B").width = 30;
    ws.column("C").hidden = true;
    ws.row(2).height = 40;
    ws.row(3).hidden = true;
    wb = await reopen(wb);
    ws = wb.get("Data");
    expect(ws.column("B").width).toBe(30);
    expect(ws.column("C").hidden).toBe(true);
    expect(ws.row(2).height).toBe(40);
    expect(ws.row(3).hidden).toBe(true);
  });
});

describe("M4 insert / delete rows and columns", () => {
  it("insertRows shifts cells and merges down", async () => {
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    let ws = wb.get("Data");
    // before: A2 = "Alice", merge A5:C5
    ws.insertRows(2, 2);
    wb = await reopen(wb);
    ws = wb.get("Data");
    expect(ws.cell("A4").value).toBe("Alice"); // row 2 -> 4
    expect(ws.mergedCells).toEqual(["A7:C7"]); // A5:C5 -> A7:C7
    expect(ws.cell("A1").value).toBe("Name"); // header unchanged
  });

  it("deleteRows removes and shifts up", async () => {
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    let ws = wb.get("Data");
    ws.deleteRows(2, 1); // remove the "Alice" row
    wb = await reopen(wb);
    ws = wb.get("Data");
    expect(ws.cell("A2").value).toBe("Bob"); // row 3 -> 2
    expect(ws.mergedCells).toEqual(["A4:C4"]); // A5:C5 -> A4:C4
  });

  it("insertColumns shifts cells right", async () => {
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    let ws = wb.get("Data");
    ws.insertColumns(1, 1); // insert a column at A
    wb = await reopen(wb);
    ws = wb.get("Data");
    expect(ws.cell("B1").value).toBe("Name"); // A1 -> B1
    expect(ws.cell("C2").value).toBe(42); // B2 -> C2
  });

  it("deleteColumns removes and shifts left", async () => {
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    let ws = wb.get("Data");
    ws.deleteColumns(1, 1); // delete column A (names)
    wb = await reopen(wb);
    ws = wb.get("Data");
    expect(ws.cell("A2").value).toBe(42); // B2 -> A2
  });
});

describe("M4 freeze panes, autofilter, tab color, title, state", () => {
  it("round-trips freeze panes and autofilter", async () => {
    let wb = await Workbook.open(fixturePath("formulas.xlsx"));
    let ws = wb.active;
    ws.freezePanes = "B2";
    ws.autoFilter = "A1:B5";
    wb = await reopen(wb);
    ws = wb.active;
    expect(ws.freezePanes).toBe("B2");
    expect(ws.autoFilter).toBe("A1:B5");
  });

  it("sets title, sheet state, and tab color", async () => {
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    const ws = wb.get("Data");
    ws.title = "Renamed";
    ws.tabColor = "FF0000";
    wb.get("Notes").sheetState = "hidden";
    wb = await reopen(wb);
    expect(wb.sheetnames).toContain("Renamed");
    expect(wb.get("Renamed").tabColor).toBe("00FF0000");
    expect(wb.get("Notes").sheetState).toBe("hidden");
  });
});

describe("M4 workbook sheet management", () => {
  it("creates, copies, moves, and removes sheets", async () => {
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    const created = wb.createSheet("Fresh");
    created.cell("A1").value = "new sheet";
    wb.copyWorksheet(wb.get("Data"), "DataCopy");
    wb = await reopen(wb);
    expect(wb.sheetnames).toContain("Fresh");
    expect(wb.sheetnames).toContain("DataCopy");
    expect(wb.get("Fresh").cell("A1").value).toBe("new sheet");
    expect(wb.get("DataCopy").cell("A1").value).toBe("Name"); // copied content

    wb.moveSheet("Fresh", 0);
    wb.removeSheet("Hidden");
    wb = await reopen(wb);
    expect(wb.sheetnames[0]).toBe("Fresh");
    expect(wb.sheetnames).not.toContain("Hidden");
  });
});
