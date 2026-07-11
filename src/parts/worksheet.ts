/** Worksheet part (xl/worksheets/sheetN.xml). */
import { OpcPackage, XmlPart } from "../opc/package.js";
import { CONTENT_TYPE as CT } from "../opc/constants.js";
import { nsdecls } from "../oxml/ns.js";
import { CT_Worksheet } from "../oxml/worksheet.js";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export class WorksheetPart extends XmlPart {
  get ctWorksheet(): CT_Worksheet {
    return new CT_Worksheet(this.root);
  }

  /** Create a new, empty worksheet part. The caller relates it to the workbook. */
  static createNew(pkg: OpcPackage): WorksheetPart {
    const partname = pkg.nextPartname("/xl/worksheets/sheet{n}.xml");
    const blob = new TextEncoder().encode(
      `${XML_DECL}<worksheet ${nsdecls("main", "r")}><sheetData/></worksheet>`,
    );
    const part = new WorksheetPart(partname, CT.SML_WORKSHEET, pkg, blob);
    pkg.markStructureDirty();
    return part;
  }

  /** Create a worksheet part that is a deep copy of `source`'s XML (structure
   * and values; the source part's own drawing/hyperlink rels are NOT cloned in
   * v1 — see the deferred list). */
  static createCopy(pkg: OpcPackage, sourceBlob: Uint8Array): WorksheetPart {
    const partname = pkg.nextPartname("/xl/worksheets/sheet{n}.xml");
    const part = new WorksheetPart(partname, CT.SML_WORKSHEET, pkg, sourceBlob);
    pkg.markStructureDirty();
    return part;
  }
}
