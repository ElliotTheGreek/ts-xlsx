import { describe, expect, it } from "vitest";
import { OpcPackage } from "../../src/opc/package.js";
import { allXlsxInputs, bytesEqual, firstDiff, zipEntries, zipEntriesOf } from "../helpers/zip.js";

/**
 * Milestone-0 acceptance: open → save with zero edits produces an equivalent
 * package — same entry set, every decompressed entry byte-identical (the zip
 * container itself is re-deflated; entry bytes are what "byte-identical" means).
 * The chart.xlsx / image.xlsx inputs specifically prove ts-xlsx preserves parts
 * openpyxl drops on its own reload.
 */
describe("zero-edit package round-trip", () => {
  for (const { name, path } of allXlsxInputs()) {
    it(name, async () => {
      const original = await zipEntriesOf(path);
      const pkg = await OpcPackage.open(path);
      const saved = await zipEntries(await pkg.toBuffer());

      const originalNames = [...original.keys()].sort();
      const savedNames = [...saved.keys()].sort();
      expect(savedNames).toEqual(originalNames);

      for (const [entryName, bytes] of original) {
        const out = saved.get(entryName)!;
        expect(bytesEqual(out, bytes), `${entryName}: ${firstDiff(out, bytes)}`).toBe(true);
      }
    });
  }

  it("preserves original entry order", async () => {
    const { path } = allXlsxInputs()[0]!;
    const original = await zipEntriesOf(path);
    const pkg = await OpcPackage.open(path);
    const saved = await zipEntries(await pkg.toBuffer());
    expect([...saved.keys()]).toEqual([...original.keys()]);
  });
});
