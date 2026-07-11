import { describe, expect, it } from "vitest";
import { Workbook, Font, PatternFill, Border, Side, Alignment, formatValue } from "../src/index.js";
import { bytesEqual, fixturePath, zipEntries, zipEntriesOf } from "./helpers/zip.js";

async function edit(fixture: string, mutate: (wb: Workbook) => void): Promise<Workbook> {
  const wb = await Workbook.open(fixturePath(fixture));
  mutate(wb);
  return Workbook.open(await wb.toBuffer());
}

describe("M3 style reads (non-mutating)", () => {
  it("reads font, fill, border, alignment, and number format from basic.xlsx", async () => {
    const wb = await Workbook.open(fixturePath("basic.xlsx"));
    const ws = wb.get("Data");

    const a1 = ws.cell("A1");
    expect(a1.font.bold).toBe(true);
    expect(a1.font.size).toBe(14);
    expect(a1.font.color?.rgb).toBe("00CC0000");
    expect(a1.alignment.horizontal).toBe("center");
    expect(a1.border.bottom.style).toBe("thin");

    const b1 = ws.cell("B1");
    const fill = b1.fill as PatternFill;
    expect(fill.patternType).toBe("solid");
    expect(fill.fgColor?.rgb).toBe("00FFFF00");

    expect(ws.cell("B2").numberFormat).toBe("#,##0.00");
    expect(ws.cell("A2").font.bold).toBeFalsy(); // default font
  });

  it("does not mutate on style reads", async () => {
    const path = fixturePath("basic.xlsx");
    const original = await zipEntriesOf(path);
    const wb = await Workbook.open(path);
    const ws = wb.get("Data");
    for (const row of ws.iterRows()) {
      for (const c of row) {
        void c.font;
        void c.fill;
        void c.border;
        void c.alignment;
        void c.numberFormat;
      }
    }
    const saved = await zipEntries(await wb.toBuffer());
    for (const [name, bytes] of original) {
      expect(bytesEqual(saved.get(name)!, bytes), `${name} changed on style read`).toBe(true);
    }
  });
});

describe("M3 style writes round-trip", () => {
  it("applies font, fill, border, alignment, and number format", async () => {
    const wb = await edit("basic.xlsx", (wb) => {
      const ws = wb.get("Data");
      ws.cell("A2").font = new Font({ bold: true, italic: true, color: "0000FF", size: 12 });
      ws.cell("B3").fill = new PatternFill({ patternType: "solid", fgColor: "FF9900" });
      ws.cell("C3").border = new Border({ top: new Side({ style: "medium" }) });
      ws.cell("A3").alignment = new Alignment({ horizontal: "right", wrapText: true });
      ws.cell("B2").numberFormat = "0.00%";
    });
    const ws = wb.get("Data");
    expect(ws.cell("A2").font.bold).toBe(true);
    expect(ws.cell("A2").font.italic).toBe(true);
    expect(ws.cell("A2").font.size).toBe(12);
    expect(ws.cell("A2").font.color?.rgb).toBe("000000FF");
    expect((ws.cell("B3").fill as PatternFill).fgColor?.rgb).toBe("00FF9900");
    expect(ws.cell("C3").border.top.style).toBe("medium");
    expect(ws.cell("A3").alignment.horizontal).toBe("right");
    expect(ws.cell("A3").alignment.wrapText).toBe(true);
    expect(ws.cell("B2").numberFormat).toBe("0.00%");
    // the underlying value survives a pure number-format change
    expect(ws.cell("B2").value).toBe(42);
  });

  it("reuses cellXfs entries instead of bloating the table", async () => {
    const before = await zipEntriesOf(fixturePath("basic.xlsx"));
    const beforeXfs = countXfs(new TextDecoder().decode(before.get("xl/styles.xml")!));

    const wb = await Workbook.open(fixturePath("basic.xlsx"));
    const ws = wb.get("Data");
    const f = new Font({ bold: true, color: "112233" });
    ws.cell("A2").font = f;
    ws.cell("A3").font = f; // identical → must reuse the same new xf
    ws.cell("A4").font = f;
    const saved = await zipEntries(await wb.toBuffer());
    const afterXfs = countXfs(new TextDecoder().decode(saved.get("xl/styles.xml")!));

    // exactly one new xf added despite three cells styled
    expect(afterXfs).toBe(beforeXfs + 1);
  });

  it("a style edit touches only styles.xml and the worksheet part", async () => {
    const original = await zipEntriesOf(fixturePath("basic.xlsx"));
    const wb = await Workbook.open(fixturePath("basic.xlsx"));
    wb.get("Data").cell("A2").font = new Font({ bold: true });
    const saved = await zipEntries(await wb.toBuffer());
    const changed = [...original.keys()]
      .filter((n) => !bytesEqual(saved.get(n)!, original.get(n)!))
      .sort();
    expect(changed).toEqual(["xl/styles.xml", "xl/worksheets/sheet1.xml"]);
  });
});

describe("M3 number-format display renderer", () => {
  it("renders numbers, currency, percent, and dates", () => {
    expect(formatValue(1234.5, "#,##0.00")).toBe("1,234.50");
    expect(formatValue(1234.5, "$#,##0.00")).toBe("$1,234.50");
    expect(formatValue(0.1234, "0.00%")).toBe("12.34%");
    expect(formatValue(1000, "#,##0")).toBe("1,000");
    expect(formatValue(46023, "yyyy-mm-dd")).toBe("2026-01-01");
    expect(formatValue(new Date(Date.UTC(2026, 0, 1)), "d-mmm-yy")).toBe("1-Jan-26");
    expect(formatValue(46214.39583333334, "yyyy-mm-dd h:mm:ss")).toBe("2026-07-11 9:30:00");
    expect(formatValue("hello", "General")).toBe("hello");
  });
});

function countXfs(stylesXml: string): number {
  const m = stylesXml.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/);
  if (m === null) return 0;
  return (m[1]!.match(/<xf[ />]/g) ?? []).length;
}
