/**
 * OPC-local element wrappers — port of pptx/opc/oxml.py.
 *
 * `.rels` and `[Content_Types].xml` use a *default* namespace, so all child
 * matching is URI-based and created children are unprefixed.
 */
import { XmlElement } from "../xml/dom.js";
import { serializeElement } from "../xml/serializer.js";
import { OxmlWrapper, xmlFragment } from "../oxml/base.js";
import { ST_TargetMode, XsdAnyUri, XsdId, XsdString } from "../oxml/simpletypes.js";
import { OPC_NAMESPACE, RELATIONSHIP_TARGET_MODE as RTM } from "./constants.js";
import { PackURI } from "./packuri.js";

/** The declaration PowerPoint itself writes; used for all regenerated/new parts. */
export const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n`;

/** Produce XML-file bytes (declaration + element) for a regenerated part. */
export function serializePartXml(root: XmlElement): Uint8Array {
  return new TextEncoder().encode(XML_DECL + serializeElement(root));
}

export class CT_Default extends OxmlWrapper {
  get extension(): string {
    return this.reqAttr("Extension", XsdString);
  }
  get contentType(): string {
    return this.reqAttr("ContentType", XsdString);
  }
}

export class CT_Override extends OxmlWrapper {
  get partName(): string {
    return this.reqAttr("PartName", XsdAnyUri);
  }
  get contentType(): string {
    return this.reqAttr("ContentType", XsdString);
  }
}

export class CT_Relationship extends OxmlWrapper {
  get rId(): string {
    return this.reqAttr("Id", XsdId);
  }
  get reltype(): string {
    return this.reqAttr("Type", XsdAnyUri);
  }
  get targetRef(): string {
    return this.reqAttr("Target", XsdAnyUri);
  }
  get targetMode(): "Internal" | "External" {
    return this.optAttrDflt("TargetMode", ST_TargetMode, RTM.INTERNAL);
  }

  static new(
    rId: string,
    reltype: string,
    targetRef: string,
    targetMode: "Internal" | "External" = RTM.INTERNAL,
  ): CT_Relationship {
    const el = new XmlElement("Relationship");
    el.setAttr("Id", rId);
    el.setAttr("Type", reltype);
    el.setAttr("Target", targetRef);
    if (targetMode === RTM.EXTERNAL) el.setAttr("TargetMode", targetMode);
    return new CT_Relationship(el);
  }
}

export class CT_Relationships extends OxmlWrapper {
  static new(): CT_Relationships {
    return new CT_Relationships(
      xmlFragment(`<Relationships xmlns="${OPC_NAMESPACE.RELATIONSHIPS}"/>`),
    );
  }

  static fromRoot(root: XmlElement): CT_Relationships {
    return new CT_Relationships(root);
  }

  get relationshipLst(): CT_Relationship[] {
    return this.el
      .findAll(OPC_NAMESPACE.RELATIONSHIPS, "Relationship")
      .map((e) => new CT_Relationship(e));
  }

  addRel(rId: string, reltype: string, target: string, isExternal = false): CT_Relationship {
    const rel = CT_Relationship.new(
      rId,
      reltype,
      target,
      isExternal ? RTM.EXTERNAL : RTM.INTERNAL,
    );
    this.el.appendChild(rel.el);
    return rel;
  }

  xmlFileBytes(): Uint8Array {
    return serializePartXml(this.el);
  }
}

export class CT_Types extends OxmlWrapper {
  static new(): CT_Types {
    return new CT_Types(xmlFragment(`<Types xmlns="${OPC_NAMESPACE.CONTENT_TYPES}"/>`));
  }

  static fromRoot(root: XmlElement): CT_Types {
    return new CT_Types(root);
  }

  get defaultLst(): CT_Default[] {
    return this.el.findAll(OPC_NAMESPACE.CONTENT_TYPES, "Default").map((e) => new CT_Default(e));
  }

  get overrideLst(): CT_Override[] {
    return this.el.findAll(OPC_NAMESPACE.CONTENT_TYPES, "Override").map((e) => new CT_Override(e));
  }

  addDefault(ext: string, contentType: string): void {
    const el = new XmlElement("Default");
    el.setAttr("Extension", ext);
    el.setAttr("ContentType", contentType);
    this.el.appendChild(el);
  }

  addOverride(partname: PackURI, contentType: string): void {
    const el = new XmlElement("Override");
    el.setAttr("PartName", partname.uri);
    el.setAttr("ContentType", contentType);
    this.el.appendChild(el);
  }

  xmlFileBytes(): Uint8Array {
    return serializePartXml(this.el);
  }
}
