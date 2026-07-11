/** Styles part (xl/styles.xml). */
import { XmlPart } from "../opc/package.js";
import { CT_Stylesheet } from "../oxml/styles.js";

export class StylesPart extends XmlPart {
  get stylesheet(): CT_Stylesheet {
    return new CT_Stylesheet(this.root);
  }
}
