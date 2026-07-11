/** Worksheet drawing part (xl/drawings/drawingN.xml) — anchors pictures/charts. */
import { OpcPackage, XmlPart } from "../opc/package.js";
import { XmlElement } from "../xml/dom.js";
import { createElement } from "../oxml/base.js";
import { nsmap } from "../oxml/ns.js";
import { CONTENT_TYPE as CT } from "../opc/constants.js";
import { getColumnLetter } from "../util.js";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

let drawingIdCounter = 1;

export class DrawingPart extends XmlPart {
  get wsDr(): XmlElement {
    return this.root;
  }

  static createNew(pkg: OpcPackage): DrawingPart {
    const partname = pkg.nextPartname("/xl/drawings/drawing{n}.xml");
    // Default-namespace spreadsheetDrawing root so appended anchors come out
    // unprefixed (matching Excel/openpyxl); a: and r: declared for blips.
    const blob = new TextEncoder().encode(
      `${XML_DECL}<wsDr xmlns="${nsmap.xdr}" xmlns:a="${nsmap.a}" xmlns:r="${nsmap.r}"/>`,
    );
    const part = new DrawingPart(partname, CT.SML_DRAWING, pkg, blob);
    pkg.markStructureDirty();
    return part;
  }

  /** Append a oneCellAnchor picture referencing `embedRId`, top-left at the
   * (0-based) col/row, sized `cx`×`cy` EMU. Returns the anchor's cell ref. */
  addPicAnchor(embedRId: string, col0: number, row0: number, cx: number, cy: number): string {
    const wsDr = this.wsDr;
    const anchor = child(wsDr, "xdr:oneCellAnchor");
    const from = child(anchor, "xdr:from");
    textChild(from, "xdr:col", String(col0));
    textChild(from, "xdr:colOff", "0");
    textChild(from, "xdr:row", String(row0));
    textChild(from, "xdr:rowOff", "0");
    const ext = child(anchor, "xdr:ext");
    ext.setAttr("cx", String(cx));
    ext.setAttr("cy", String(cy));

    const pic = child(anchor, "xdr:pic");
    const nvPicPr = child(pic, "xdr:nvPicPr");
    const cNvPr = child(nvPicPr, "xdr:cNvPr");
    const id = drawingIdCounter++;
    cNvPr.setAttr("id", String(id));
    cNvPr.setAttr("name", `Image ${id}`);
    child(nvPicPr, "xdr:cNvPicPr");

    const blipFill = child(pic, "xdr:blipFill");
    const blip = child(blipFill, "a:blip");
    blip.setAttr("r:embed", embedRId);
    const stretch = child(blipFill, "a:stretch");
    child(stretch, "a:fillRect");

    const spPr = child(pic, "xdr:spPr");
    child(spPr, "a:prstGeom").setAttr("prst", "rect");

    child(anchor, "xdr:clientData");
    return `${getColumnLetter(col0 + 1)}${row0 + 1}`;
  }
}

function child(parent: XmlElement, tag: Parameters<typeof createElement>[0]): XmlElement {
  const el = createElement(tag, parent);
  parent.appendChild(el);
  return el;
}

function textChild(parent: XmlElement, tag: Parameters<typeof createElement>[0], text: string): XmlElement {
  const el = child(parent, tag);
  el.setText(text);
  return el;
}
