import { describe, expect, it } from "vitest";
import { parseXml } from "../../src/xml/parser.js";
import { serializeDocument } from "../../src/xml/serializer.js";
import { allXlsxInputs, bytesEqual, firstDiff, zipEntriesOf } from "../helpers/zip.js";

/**
 * The backbone property test: for every XML entry of every available .xlsx,
 * serialize(parse(bytes)) === bytes. This is what makes "touched parts differ
 * only at the edited nodes" provable for the whole library.
 */
describe("XML byte round-trip over every xlsx entry", () => {
  const inputs = allXlsxInputs();
  it("has xlsx inputs to test", () => {
    expect(inputs.length).toBeGreaterThan(0);
  });

  for (const { name, path } of inputs) {
    it(name, async () => {
      const entries = await zipEntriesOf(path);
      let xmlCount = 0;
      for (const [entryName, bytes] of entries) {
        if (!/\.(xml|rels)$/i.test(entryName)) continue;
        xmlCount++;
        const out = serializeDocument(parseXml(bytes));
        expect(bytesEqual(out, bytes), `${entryName}: ${firstDiff(out, bytes)}`).toBe(true);
      }
      expect(xmlCount).toBeGreaterThan(0);
    });
  }
});
