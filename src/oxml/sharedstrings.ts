/**
 * Shared string table wrappers — port of openpyxl/cell/text.py + the sst reader.
 * `<sst>` holds `<si>` string items; each item is either a single `<t>` or a
 * run of `<r>` rich-text pieces (each with its own `<t>`). Phonetic `<rPh>`
 * guides are ignored for text extraction (they are not `<r>` runs).
 */
import { OxmlWrapper, createElement } from "./base.js";
import { nsmap } from "./ns.js";

const MAIN = nsmap.main;

/** `<si>` shared string item. */
export class CT_Si extends OxmlWrapper {
  /** Plain concatenated text of the item (rich-run text joined, phonetics skipped). */
  get text(): string {
    const directT = this.el.find(MAIN, "t");
    if (directT !== null) return directT.text;
    let out = "";
    for (const r of this.el.findAll(MAIN, "r")) {
      const t = r.find(MAIN, "t");
      if (t !== null) out += t.text;
    }
    return out;
  }
}

/** `<sst>` shared string table root. */
export class CT_Sst extends OxmlWrapper {
  get siLst(): CT_Si[] {
    return this.el.findAll(MAIN, "si").map((e) => new CT_Si(e));
  }

  /** Number of `<si>` entries. */
  get count(): number {
    return this.el.findAll(MAIN, "si").length;
  }

  /** Text of the `<si>` at `index`, or undefined when out of range. */
  textAt(index: number): string | undefined {
    const items = this.el.findAll(MAIN, "si");
    const si = items[index];
    return si === undefined ? undefined : new CT_Si(si).text;
  }

  /** Append a plain-text `<si>` and return its index (M2 write path). */
  addPlain(text: string): number {
    const index = this.count;
    const si = this.addChild("main:si");
    const t = createElement("main:t", si);
    si.appendChild(t);
    t.setText(text);
    if (text.trim().length !== text.length) t.setAttr("xml:space", "preserve");
    // count / uniqueCount are Excel hints; keep them consistent with the table.
    const n = this.count;
    this.el.setAttr("count", String(n));
    this.el.setAttr("uniqueCount", String(n));
    return index;
  }
}
