import { describe, expect, it } from "vitest";
import { Workbook } from "../src/index.js";
import { bytesEqual, fixturePath, zipEntries, zipEntriesOf } from "./helpers/zip.js";

/** Edit a workbook in memory and reopen the saved bytes. */
async function edit(fixture: string, mutate: (wb: Workbook) => void): Promise<Workbook> {
  const wb = await Workbook.open(fixturePath(fixture));
  mutate(wb);
  const bytes = await wb.toBuffer();
  return Workbook.open(bytes);
}

describe("M2 cell writes round-trip through ts-xlsx", () => {
  it("writes numbers, booleans, strings, and formulas", async () => {
    const wb = await edit("basic.xlsx", (wb) => {
      const ws = wb.get("Data");
      ws.cell("B2").value = 99;
      ws.cell("D1").value = true;
      ws.cell("D2").value = "hello world";
      ws.cell("D3").value = "=B2*2";
    });
    const ws = wb.get("Data");
    expect(ws.cell("B2").value).toBe(99);
    expect(ws.cell("D1").value).toBe(true);
    expect(ws.cell("D2").value).toBe("hello world");
    expect(ws.cell("D3").value).toBe("=B2*2");
    expect(ws.cell("D3").dataType).toBe("f");
  });

  it("writes a Date and reads it back as a Date (date number format applied)", async () => {
    const when = new Date(Date.UTC(2030, 5, 15, 13, 45, 0));
    const wb = await edit("basic.xlsx", (wb) => {
      wb.get("Data").cell("E1").value = when;
    });
    const got = wb.get("Data").cell("E1").value as Date;
    expect(got).toBeInstanceOf(Date);
    expect([got.getUTCFullYear(), got.getUTCMonth(), got.getUTCDate()]).toEqual([2030, 5, 15]);
    expect(got.getUTCHours()).toBe(13);
    expect(got.getUTCMinutes()).toBe(45);
  });

  it("clears a cell when assigned null", async () => {
    const wb = await edit("basic.xlsx", (wb) => {
      wb.get("Data").cell("A1").value = null;
    });
    expect(wb.get("Data").cell("A1").value).toBeNull();
  });

  it("append() adds a row at maxRow+1", async () => {
    const wb = await edit("basic.xlsx", (wb) => {
      wb.get("Data").append(["Carol", 88, false]);
    });
    const ws = wb.get("Data");
    expect(ws.cell("A6").value).toBe("Carol");
    expect(ws.cell("B6").value).toBe(88);
    expect(ws.cell("C6").value).toBe(false);
  });
});

describe("M2 shared string interning", () => {
  it("creates a sharedStrings part when absent and dedupes repeats", async () => {
    // types.xlsx (openpyxl) has NO sharedStrings part — writing strings creates it
    const wb = await Workbook.open(fixturePath("types.xlsx"));
    const ws = wb.active;
    ws.cell("C1").value = "same";
    ws.cell("C2").value = "same";
    ws.cell("C3").value = "different";
    const bytes = await wb.toBuffer();

    const entries = await zipEntries(bytes);
    expect(entries.has("xl/sharedStrings.xml")).toBe(true);
    const sst = new TextDecoder().decode(entries.get("xl/sharedStrings.xml")!);
    // exactly two distinct <si> entries ("same", "different")
    expect((sst.match(/<si>/g) ?? []).length).toBe(2);

    const reopened = await Workbook.open(bytes);
    expect(reopened.active.cell("C1").value).toBe("same");
    expect(reopened.active.cell("C2").value).toBe("same");
    expect(reopened.active.cell("C3").value).toBe("different");
  });
});

describe("M2 fidelity — an edit touches only what it must", () => {
  it("a number edit changes only the one worksheet part", async () => {
    const original = await zipEntriesOf(fixturePath("basic.xlsx"));
    const wb = await Workbook.open(fixturePath("basic.xlsx"));
    wb.get("Data").cell("B2").value = 99;
    const saved = await zipEntries(await wb.toBuffer());

    const changed = [...original.keys()].filter(
      (name) => !bytesEqual(saved.get(name)!, original.get(name)!),
    );
    expect(changed).toEqual(["xl/worksheets/sheet1.xml"]);
    expect([...saved.keys()].sort()).toEqual([...original.keys()].sort());
  });

  it("a formula edit sets fullCalcOnLoad and drops the stale calcChain", async () => {
    // shared_strings.xlsx has a formula but no calcChain; editing sets calcPr
    const original = await zipEntriesOf(fixturePath("shared_strings.xlsx"));
    const wb = await Workbook.open(fixturePath("shared_strings.xlsx"));
    wb.active.cell("A3").value = "=B2*10";
    const saved = await zipEntries(await wb.toBuffer());

    const changed = [...original.keys()]
      .filter((name) => !bytesEqual(saved.get(name)!, original.get(name)!))
      .sort();
    expect(changed).toEqual(["xl/workbook.xml", "xl/worksheets/sheet1.xml"]);
    const wbXml = new TextDecoder().decode(saved.get("xl/workbook.xml")!);
    expect(wbXml).toContain('fullCalcOnLoad="1"');
  });
});
