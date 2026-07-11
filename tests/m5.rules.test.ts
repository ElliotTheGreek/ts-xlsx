import { describe, expect, it } from "vitest";
import { Workbook, PatternFill } from "../src/index.js";
import { bytesEqual, fixturePath, zipEntries, zipEntriesOf } from "./helpers/zip.js";

async function reopen(wb: Workbook): Promise<Workbook> {
  return Workbook.open(await wb.toBuffer());
}

describe("M5 reads (conditional formatting, validation, hyperlinks, names, comments)", () => {
  it("reads conditional formatting rules", async () => {
    const wb = await Workbook.open(fixturePath("rules.xlsx"));
    const cf = wb.get("Sheet1").conditionalFormatting;
    expect(cf).toHaveLength(2);
    const cellIs = cf.find((c) => c.sqref === "A1:A10")!;
    expect(cellIs.rules[0]!.type).toBe("cellIs");
    expect(cellIs.rules[0]!.operator).toBe("greaterThan");
    expect(cellIs.rules[0]!.formulas).toEqual(["5"]);
    expect(cf.find((c) => c.sqref === "B1:B10")!.rules[0]!.type).toBe("colorScale");
  });

  it("reads data validations", async () => {
    const wb = await Workbook.open(fixturePath("rules.xlsx"));
    const dv = wb.get("Sheet1").dataValidations;
    expect(dv).toHaveLength(1);
    expect(dv[0]!.type).toBe("list");
    expect(dv[0]!.sqref).toBe("C1:C5");
    expect(dv[0]!.formula1).toBe('"apple,banana,cherry"');
  });

  it("reads hyperlinks and defined names", async () => {
    const wb = await Workbook.open(fixturePath("rules.xlsx"));
    const links = wb.get("Sheet1").hyperlinks;
    expect(links).toHaveLength(1);
    expect(links[0]!.ref).toBe("D1");
    expect(links[0]!.target).toBe("https://flowdot.ai");
    expect(wb.get("Sheet1").cell("D1").hyperlink).toBe("https://flowdot.ai");
    expect(wb.definedNames.get("MyRange")).toBe("Sheet1!$A$1:$A$10");
  });

  it("reads cell comments", async () => {
    const wb = await Workbook.open(fixturePath("comments.xlsx"));
    const comments = wb.active.comments;
    expect(comments.length).toBeGreaterThan(0);
    expect(comments[0]!.text).toContain("a note");
    expect(wb.active.cell(comments[0]!.ref).comment?.text).toContain("a note");
  });
});

describe("M5 writes round-trip", () => {
  it("adds a cellIs rule with a highlight fill", async () => {
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    wb.get("Data").addCellIsRule("B2:B3", {
      operator: "greaterThan",
      formula: "20",
      fill: new PatternFill({ patternType: "solid", fgColor: "FFC7CE" }),
    });
    wb = await reopen(wb);
    const cf = wb.get("Data").conditionalFormatting;
    expect(cf).toHaveLength(1);
    expect(cf[0]!.sqref).toBe("B2:B3");
    expect(cf[0]!.rules[0]!.operator).toBe("greaterThan");
    expect(cf[0]!.rules[0]!.dxfId).toBe(0);
  });

  it("adds a list data validation and a hyperlink", async () => {
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    const ws = wb.get("Data");
    ws.addDataValidation({ sqref: "D1:D5", type: "list", formula1: '"x,y,z"', allowBlank: true });
    ws.cell("A1").hyperlink = "https://example.com";
    wb = await reopen(wb);
    const dv = wb.get("Data").dataValidations;
    expect(dv[0]!.type).toBe("list");
    expect(dv[0]!.formula1).toBe('"x,y,z"');
    expect(wb.get("Data").cell("A1").hyperlink).toBe("https://example.com");
  });

  it("sets and removes a defined name", async () => {
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    wb.setDefinedName("Prices", "Data!$B$2:$B$3");
    wb = await reopen(wb);
    expect(wb.definedNames.get("Prices")).toBe("Data!$B$2:$B$3");
    wb.removeDefinedName("Prices");
    wb = await reopen(wb);
    expect(wb.definedNames.has("Prices")).toBe(false);
  });
});

describe("M5 fidelity — targeted edits", () => {
  it("a defined-name edit touches only workbook.xml", async () => {
    const original = await zipEntriesOf(fixturePath("basic.xlsx"));
    const wb = await Workbook.open(fixturePath("basic.xlsx"));
    wb.setDefinedName("Foo", "Data!$A$1");
    const saved = await zipEntries(await wb.toBuffer());
    const changed = [...original.keys()]
      .filter((n) => !bytesEqual(saved.get(n)!, original.get(n)!))
      .sort();
    expect(changed).toEqual(["xl/workbook.xml"]);
  });
});
