import { XmlCdata, XmlComment, XmlDocument, XmlElement, XmlNode, XmlPI, XmlText } from "./dom.js";

/** Serialize a document back to UTF-8 bytes (byte-identical when unmodified). */
export function serializeDocument(doc: XmlDocument): Uint8Array {
  return new TextEncoder().encode(serializeDocumentToString(doc));
}

export function serializeDocumentToString(doc: XmlDocument): string {
  const parts: string[] = [];
  if (doc.hadBom) parts.push("﻿");
  if (doc.declRaw !== null) parts.push(doc.declRaw);
  for (const n of doc.prolog) writeNode(n, parts);
  writeNode(doc.root, parts);
  for (const n of doc.epilog) writeNode(n, parts);
  return parts.join("");
}

export function serializeElement(el: XmlElement): string {
  const parts: string[] = [];
  writeNode(el, parts);
  return parts.join("");
}

function writeNode(node: XmlNode, out: string[]): void {
  if (node instanceof XmlElement) {
    out.push("<", node.name);
    for (const a of node.attrs) {
      out.push(a.leadingGap, a.name, "=", a.quote, a.rawValue, a.quote);
    }
    if (node.children.length === 0 && node.selfClosing) {
      out.push(node.endTagGap, "/>");
      return;
    }
    out.push(node.endTagGap, ">");
    for (const c of node.children) writeNode(c, out);
    out.push("</", node.name, ">");
    return;
  }
  if (node instanceof XmlText) {
    out.push(node.raw);
    return;
  }
  if (node instanceof XmlCdata) {
    out.push("<![CDATA[", node.value, "]]>");
    return;
  }
  if (node instanceof XmlComment) {
    out.push("<!--", node.text, "-->");
    return;
  }
  out.push((node as XmlPI).raw);
}
