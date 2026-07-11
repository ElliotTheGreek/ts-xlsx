/** Main workbook part (xl/workbook.xml) — the office-document part. */
import { XmlPart } from "../opc/package.js";
import { Workbook } from "../workbook.js";

export class WorkbookPart extends XmlPart {
  #workbook: Workbook | undefined;

  get workbook(): Workbook {
    if (this.#workbook === undefined) {
      this.#workbook = new Workbook(this);
    }
    return this.#workbook;
  }
}
