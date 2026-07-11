/** Shared strings part (xl/sharedStrings.xml). */
import { OpcPackage, XmlPart } from "../opc/package.js";
import { PackURI } from "../opc/packuri.js";
import { CONTENT_TYPE as CT } from "../opc/constants.js";
import { nsdecls } from "../oxml/ns.js";
import { CT_Sst } from "../oxml/sharedstrings.js";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export class SharedStringsPart extends XmlPart {
  get sst(): CT_Sst {
    return new CT_Sst(this.root);
  }

  /** Create an empty shared-strings part (for workbooks that lack one — e.g.
   * openpyxl output, which writes inline strings). The caller relates it to
   * the workbook part. */
  static createNew(pkg: OpcPackage): SharedStringsPart {
    const partname = new PackURI("/xl/sharedStrings.xml");
    const blob = new TextEncoder().encode(`${XML_DECL}<sst ${nsdecls("main")} count="0" uniqueCount="0"/>`);
    const part = new SharedStringsPart(partname, CT.SML_SHARED_STRINGS, pkg, blob);
    pkg.markStructureDirty();
    return part;
  }
}
