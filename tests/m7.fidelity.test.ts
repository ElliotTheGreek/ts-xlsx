import { describe, expect, it } from "vitest";
import { Workbook } from "../src/index.js";
import { OpcPackage } from "../src/opc/package.js";
import { parseXml } from "../src/xml/parser.js";
import { serializeDocument } from "../src/xml/serializer.js";
import { allXlsxInputs, bytesEqual, fixturePath, zipEntries, zipEntriesOf } from "./helpers/zip.js";

/**
 * M7 acceptance: the fidelity guarantees, stated as explicit invariants over the
 * whole corpus rather than a single fixture.
 */
describe("M7 backbone — serialize(parse(bytes)) === bytes for every XML entry", () => {
  it("holds across the entire fixture corpus", async () => {
    let total = 0;
    for (const { path } of allXlsxInputs()) {
      for (const [name, bytes] of await zipEntriesOf(path)) {
        if (!/\.(xml|rels)$/i.test(name)) continue;
        total++;
        expect(bytesEqual(serializeDocument(parseXml(bytes)), bytes), `${path} :: ${name}`).toBe(true);
      }
    }
    expect(total).toBeGreaterThan(30);
  });
});

describe("M7 orphan-entry preservation (openpyxl drops these; ts-xlsx must not)", () => {
  it("preserves zip entries unreachable from the relationship graph", async () => {
    const path = fixturePath("orphan.xlsx");
    const original = await zipEntriesOf(path);
    expect(original.has("xl/vendorData/custom.xml")).toBe(true);
    expect(original.has("customNotes.txt")).toBe(true);

    const pkg = await OpcPackage.open(path);
    const saved = await zipEntries(await pkg.toBuffer());

    expect(saved.has("xl/vendorData/custom.xml")).toBe(true);
    expect(saved.has("customNotes.txt")).toBe(true);
    expect(bytesEqual(saved.get("xl/vendorData/custom.xml")!, original.get("xl/vendorData/custom.xml")!)).toBe(true);
    expect(bytesEqual(saved.get("customNotes.txt")!, original.get("customNotes.txt")!)).toBe(true);
  });

  it("survives an edit through the public Workbook API too", async () => {
    const original = await zipEntriesOf(fixturePath("orphan.xlsx"));
    const wb = await Workbook.open(fixturePath("orphan.xlsx"));
    wb.get("Data").cell("B2").value = 123;
    const saved = await zipEntries(await wb.toBuffer());
    expect(bytesEqual(saved.get("customNotes.txt")!, original.get("customNotes.txt")!)).toBe(true);
  });
});

describe("M7 reads never mutate — corpus-wide", () => {
  it("open → read the whole surface → save is byte-identical for every fixture", async () => {
    for (const { name, path } of allXlsxInputs()) {
      const original = await zipEntriesOf(path);
      const wb = await Workbook.open(path);
      for (const ws of wb.worksheets) {
        void ws.dimensions;
        void ws.mergedCells;
        void ws.conditionalFormatting;
        void ws.hyperlinks;
        void ws.charts;
        void ws.images;
        for (const row of ws.iterRows()) {
          for (const c of row) {
            void c.value;
            void c.font;
            void c.numberFormat;
          }
        }
      }
      void wb.definedNames;
      const saved = await zipEntries(await wb.toBuffer());
      for (const [entry, bytes] of original) {
        expect(bytesEqual(saved.get(entry)!, bytes), `${name} :: ${entry} mutated on read`).toBe(true);
      }
    }
  });
});
