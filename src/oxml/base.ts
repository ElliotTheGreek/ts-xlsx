/**
 * OxmlWrapper — the xmlchemy replacement (pptx/oxml/xmlchemy.py).
 *
 * python-pptx injects attribute/child accessors onto element classes with a
 * metaclass. Here, wrapper classes are plain TS classes constructed on demand
 * around raw XmlElements; the xmlchemy vocabulary (RequiredAttribute,
 * OptionalAttribute, ZeroOrOne, OneAndOnlyOne, ZeroOrMore + insert-before-
 * successors) is provided as protected helpers.
 */
import { InvalidXmlError } from "../exc.js";
import { XmlElement, XmlText } from "../xml/dom.js";
import { parseXml } from "../xml/parser.js";
import { NsTag, nsmap, splitNsTag } from "./ns.js";
import { SimpleType } from "./simpletypes.js";

/** Parse an XML string into a detached element (python's parse_xml on templates). */
export function xmlFragment(xml: string): XmlElement {
  const doc = parseXml(xml);
  const root = doc.root;
  root.docIfRoot = null; // detach from the throwaway document
  return root;
}

/**
 * Create an element for insertion into an existing tree — the replacement for
 * lxml's automatic prefix management. Reuses an in-scope prefix of `context`
 * when the URI is already declared (including the default namespace);
 * otherwise uses the canonical prefix and declares it on the new element.
 */
export function createElement(tag: NsTag, context: XmlElement): XmlElement {
  const { nsUri, localName, prefix } = splitNsTag(tag);
  const inScope = context.lookupPrefixFor(nsUri);
  if (inScope === "") return new XmlElement(localName);
  if (typeof inScope === "string") return new XmlElement(`${inScope}:${localName}`);
  const el = new XmlElement(`${prefix}:${localName}`);
  el.setAttr(`xmlns:${prefix}`, nsUri);
  return el;
}

/** Count occurrences of an attribute value anywhere in a subtree — replaces
 * python-pptx's `//@r:id` xpath in drop_rel's reference counting. */
export function countAttrValues(
  root: XmlElement,
  nsUri: string,
  localName: string,
  value: string,
): number {
  let n = root.getAttrNS(nsUri, localName) === value ? 1 : 0;
  for (const c of root.childElements) n += countAttrValues(c, nsUri, localName, value);
  return n;
}

export abstract class OxmlWrapper {
  constructor(readonly el: XmlElement) {}

  // -- attributes (RequiredAttribute / OptionalAttribute) ----------------

  protected reqAttr<T>(name: string, t: SimpleType<T>): T {
    const raw = this.el.getAttr(name);
    if (raw === null) {
      throw new InvalidXmlError(`required attribute ${name} missing on <${this.el.name}>`);
    }
    return t.fromXml(raw);
  }

  protected optAttr<T>(name: string, t: SimpleType<T>): T | undefined {
    const raw = this.el.getAttr(name);
    return raw === null ? undefined : t.fromXml(raw);
  }

  protected optAttrDflt<T>(name: string, t: SimpleType<T>, dflt: T): T {
    return this.optAttr(name, t) ?? dflt;
  }

  protected setAttrVal<T>(name: string, t: SimpleType<T>, v: T): void {
    this.el.setAttr(name, t.toXml(v));
  }

  /** OptionalAttribute setter semantics: undefined (or the default) removes. */
  protected setOptAttrVal<T>(name: string, t: SimpleType<T>, v: T | undefined, dflt?: T): void {
    if (v === undefined || v === dflt) this.el.removeAttr(name);
    else this.el.setAttr(name, t.toXml(v));
  }

  // -- children (ZeroOrOne / OneAndOnlyOne / ZeroOrMore) -----------------

  protected zeroOrOne(tag: NsTag): XmlElement | null {
    const { nsUri, localName } = splitNsTag(tag);
    return this.el.find(nsUri, localName);
  }

  protected oneAndOnlyOne(tag: NsTag): XmlElement {
    const found = this.zeroOrOne(tag);
    if (found === null) {
      throw new InvalidXmlError(`required child <${tag}> missing on <${this.el.name}>`);
    }
    return found;
  }

  protected zeroOrMore(tag: NsTag): XmlElement[] {
    const { nsUri, localName } = splitNsTag(tag);
    return this.el.findAll(nsUri, localName);
  }

  /** Port of `get_or_add_x`. */
  protected getOrAdd(tag: NsTag, successors: readonly NsTag[]): XmlElement {
    return this.zeroOrOne(tag) ?? this.addChild(tag, successors);
  }

  /**
   * Port of `_add_x`/`_insert_x` + `insert_element_before`: insert a new
   * child before the first element matching any successor tag, else append
   * (before a trailing whitespace text node, to keep formatted files tidy).
   */
  protected addChild(tag: NsTag, successors: readonly NsTag[] = []): XmlElement {
    const child = createElement(tag, this.el);
    this.insertElementBefore(child, successors);
    return child;
  }

  protected insertElementBefore(child: XmlElement, successors: readonly NsTag[]): void {
    for (const succ of successors) {
      const { nsUri, localName } = splitNsTag(succ);
      const ref = this.el.find(nsUri, localName);
      if (ref) {
        this.el.insertBefore(child, ref);
        return;
      }
    }
    const kids = this.el.children;
    const last = kids[kids.length - 1];
    if (last instanceof XmlText && last.isWhitespace) {
      this.el.insertBefore(child, last);
    } else {
      this.el.appendChild(child);
    }
  }

  /** Port of `_remove_x` / remove_all. */
  protected removeAllChildren(...tags: NsTag[]): void {
    for (const tag of tags) {
      for (const el of this.zeroOrMore(tag)) this.el.removeChild(el);
    }
  }
}

/** Canonical URI → prefix reverse map (python's pfxmap). */
export const pfxmap: Record<string, string> = Object.fromEntries(
  Object.entries(nsmap).map(([pfx, uri]) => [uri, pfx]),
);
