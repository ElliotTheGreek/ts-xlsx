/**
 * Minimal lossless XML DOM for OOXML parts.
 *
 * Fidelity contract: `serialize(parse(bytes)) === bytes` for any well-formed
 * UTF-8 document without DOCTYPE. Everything positional is stored as written:
 * attribute order, quote characters, still-escaped raw values, inter-attribute
 * whitespace, tag shape (self-closing vs open/close), text nodes verbatim,
 * comments/PIs/CDATA, the XML declaration, and a UTF-8 BOM flag.
 *
 * Mutation contract: every mutator marks the owning document dirty; reads
 * never mutate anything. Child and attribute arrays are private — all changes
 * funnel through the methods below, so dirty tracking cannot be bypassed.
 */
import { decodeXmlText, escapeXmlAttr, escapeXmlText } from "./escape.js";

export interface XmlAttr {
  /** Attribute name exactly as written, e.g. "r:id", "xmlns", "Id". */
  readonly name: string;
  /** Value exactly as written (still escaped), without quotes. */
  readonly rawValue: string;
  readonly quote: '"' | "'";
  /** Whitespace preceding the attribute in the source (>= 1 char). */
  readonly leadingGap: string;
}

export type XmlNode = XmlElement | XmlText | XmlCdata | XmlComment | XmlPI;

abstract class XmlNodeBase {
  parent: XmlElement | null = null;

  /** The owning document, reached by walking to the root element. */
  get doc(): XmlDocument | null {
    let el: XmlElement | null = this instanceof XmlElement ? this : this.parent;
    while (el && el.parent) el = el.parent;
    return el?.docIfRoot ?? null;
  }

  protected markDirty(): void {
    this.doc?.markDirty();
  }
}

export class XmlText extends XmlNodeBase {
  /** Text exactly as written in the source (still escaped). */
  #raw: string;

  constructor(raw: string) {
    super();
    this.#raw = raw;
  }

  static fromValue(value: string): XmlText {
    return new XmlText(escapeXmlText(value));
  }

  get raw(): string {
    return this.#raw;
  }

  get value(): string {
    return decodeXmlText(this.#raw);
  }

  setValue(value: string): void {
    this.#raw = escapeXmlText(value);
    this.markDirty();
  }

  /** True when the node is only XML whitespace (inter-element formatting). */
  get isWhitespace(): boolean {
    return /^[ \t\r\n]*$/.test(this.#raw);
  }
}

export class XmlCdata extends XmlNodeBase {
  #value: string;

  constructor(value: string) {
    super();
    this.#value = value;
  }

  get value(): string {
    return this.#value;
  }
}

export class XmlComment extends XmlNodeBase {
  constructor(readonly text: string) {
    super();
  }
}

export class XmlPI extends XmlNodeBase {
  /** Full source slice including `<?` and `?>`. */
  constructor(readonly raw: string, readonly target: string) {
    super();
  }
}

export class XmlElement extends XmlNodeBase {
  /** Element name exactly as written, e.g. "p:sp" or "Relationship". */
  readonly name: string;
  /** Tag shape in the source; ignored once the element has children. */
  selfClosing: boolean;
  /** Whitespace between the last attribute (or name) and `>` / `/>`. */
  endTagGap: string;
  /** Set only on a document's root element. */
  docIfRoot: XmlDocument | null = null;

  #attrs: XmlAttr[] = [];
  #children: XmlNode[] = [];

  constructor(name: string, opts?: { selfClosing?: boolean; endTagGap?: string }) {
    super();
    this.name = name;
    this.selfClosing = opts?.selfClosing ?? true;
    this.endTagGap = opts?.endTagGap ?? "";
  }

  get localName(): string {
    const c = this.name.indexOf(":");
    return c === -1 ? this.name : this.name.slice(c + 1);
  }

  get prefix(): string | null {
    const c = this.name.indexOf(":");
    return c === -1 ? null : this.name.slice(0, c);
  }

  /** Namespace URI resolved from in-scope declarations; null if undeclared. */
  get nsUri(): string | null {
    return this.resolvePrefix(this.prefix);
  }

  /**
   * Resolve a namespace prefix (null = default namespace) against the
   * declarations in scope at this element.
   */
  resolvePrefix(prefix: string | null): string | null {
    const attrName = prefix === null ? "xmlns" : `xmlns:${prefix}`;
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    let el: XmlElement | null = this;
    while (el) {
      const v = el.getAttr(attrName);
      if (v !== null) return v === "" ? null : v;
      el = el.parent;
    }
    return null;
  }

  /** Find an in-scope prefix for a namespace URI. Returns "" for the default
   * namespace, the prefix string otherwise, or undefined when undeclared. */
  lookupPrefixFor(nsUri: string): string | "" | undefined {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    let el: XmlElement | null = this;
    const shadowed = new Set<string>();
    while (el) {
      for (const a of el.attrs) {
        if (a.name === "xmlns") {
          if (!shadowed.has("") && decodeXmlText(a.rawValue) === nsUri) return "";
          shadowed.add("");
        } else if (a.name.startsWith("xmlns:")) {
          const pfx = a.name.slice(6);
          if (!shadowed.has(pfx) && decodeXmlText(a.rawValue) === nsUri) return pfx;
          shadowed.add(pfx);
        }
      }
      el = el.parent;
    }
    return undefined;
  }

  // -- attributes ------------------------------------------------------

  get attrs(): ReadonlyArray<XmlAttr> {
    return this.#attrs;
  }

  /** Decoded value of the attribute with this literal (as-written) name. */
  getAttr(name: string): string | null {
    for (const a of this.#attrs) {
      if (a.name === name) return decodeXmlText(a.rawValue);
    }
    return null;
  }

  /** Decoded value of the attribute with this namespace URI + local name.
   * Unprefixed attributes have no namespace (pass nsUri = null for those). */
  getAttrNS(nsUri: string | null, localName: string): string | null {
    for (const a of this.#attrs) {
      const c = a.name.indexOf(":");
      if (c === -1) {
        if (nsUri === null && a.name === localName) return decodeXmlText(a.rawValue);
      } else {
        if (a.name.slice(c + 1) !== localName) continue;
        const pfx = a.name.slice(0, c);
        if (pfx === "xmlns") continue;
        if (nsUri !== null && this.resolvePrefix(pfx) === nsUri) return decodeXmlText(a.rawValue);
      }
    }
    return null;
  }

  /** Set (or append) an attribute by literal name. Marks the document dirty. */
  setAttr(name: string, value: string): void {
    for (let i = 0; i < this.#attrs.length; i++) {
      const a = this.#attrs[i]!;
      if (a.name === name) {
        this.#attrs[i] = {
          name,
          rawValue: escapeXmlAttr(value, a.quote),
          quote: a.quote,
          leadingGap: a.leadingGap,
        };
        this.markDirty();
        return;
      }
    }
    this.#attrs.push({ name, rawValue: escapeXmlAttr(value, '"'), quote: '"', leadingGap: " " });
    this.markDirty();
  }

  removeAttr(name: string): void {
    const idx = this.#attrs.findIndex((a) => a.name === name);
    if (idx === -1) return;
    this.#attrs.splice(idx, 1);
    this.markDirty();
  }

  /** Parser-only: append a fully-specified attribute without dirtying. */
  loadAttr(attr: XmlAttr): void {
    this.#attrs.push(attr);
  }

  // -- children --------------------------------------------------------

  get children(): ReadonlyArray<XmlNode> {
    return this.#children;
  }

  get childElements(): XmlElement[] {
    return this.#children.filter((n): n is XmlElement => n instanceof XmlElement);
  }

  appendChild<T extends XmlNode>(node: T): T {
    detach(node);
    node.parent = this;
    this.#children.push(node);
    this.markDirty();
    return node;
  }

  insertBefore<T extends XmlNode>(node: T, ref: XmlNode | null): T {
    if (ref === null) return this.appendChild(node);
    const idx = this.#children.indexOf(ref);
    if (idx === -1) throw new Error("insertBefore: reference node is not a child");
    detach(node);
    node.parent = this;
    this.#children.splice(idx, 0, node);
    this.markDirty();
    return node;
  }

  removeChild(node: XmlNode): void {
    const idx = this.#children.indexOf(node);
    if (idx === -1) throw new Error("removeChild: node is not a child");
    this.#children.splice(idx, 1);
    node.parent = null;
    this.markDirty();
  }

  /** Parser-only: append without dirtying. */
  loadChild(node: XmlNode): void {
    node.parent = this;
    this.#children.push(node);
  }

  // -- navigation ------------------------------------------------------

  matches(nsUri: string | null, localName: string): boolean {
    return this.localName === localName && this.nsUri === nsUri;
  }

  /** First direct child element matching nsUri + localName. */
  find(nsUri: string | null, localName: string): XmlElement | null {
    for (const c of this.#children) {
      if (c instanceof XmlElement && c.matches(nsUri, localName)) return c;
    }
    return null;
  }

  findAll(nsUri: string | null, localName: string): XmlElement[] {
    return this.childElements.filter((c) => c.matches(nsUri, localName));
  }

  /** All descendant elements matching nsUri + localName, document order. */
  findAllDeep(nsUri: string | null, localName: string): XmlElement[] {
    const out: XmlElement[] = [];
    const walk = (el: XmlElement): void => {
      for (const c of el.childElements) {
        if (c.matches(nsUri, localName)) out.push(c);
        walk(c);
      }
    };
    walk(this);
    return out;
  }

  /** Concatenated decoded text of direct text/CDATA children. */
  get text(): string {
    let out = "";
    for (const c of this.#children) {
      if (c instanceof XmlText) out += c.value;
      else if (c instanceof XmlCdata) out += c.value;
    }
    return out;
  }

  /** Replace all children with a single text node. */
  setText(value: string): void {
    for (const c of this.#children) c.parent = null;
    this.#children.length = 0;
    if (value !== "") this.loadChild(XmlText.fromValue(value));
    this.markDirty();
  }

  /** Deep copy, detached (no parent, no doc). */
  cloneNode(): XmlElement {
    const copy = new XmlElement(this.name, {
      selfClosing: this.selfClosing,
      endTagGap: this.endTagGap,
    });
    for (const a of this.#attrs) copy.loadAttr({ ...a });
    for (const c of this.#children) {
      if (c instanceof XmlElement) copy.loadChild(c.cloneNode());
      else if (c instanceof XmlText) copy.loadChild(new XmlText(c.raw));
      else if (c instanceof XmlCdata) copy.loadChild(new XmlCdata(c.value));
      else if (c instanceof XmlComment) copy.loadChild(new XmlComment(c.text));
      else copy.loadChild(new XmlPI((c as XmlPI).raw, (c as XmlPI).target));
    }
    return copy;
  }
}

function detach(node: XmlNode): void {
  if (node.parent) node.parent.removeChild(node);
}

export class XmlDocument {
  /** The XML declaration exactly as written (incl. `<?xml` and `?>`), or null. */
  declRaw: string | null = null;
  hadBom = false;
  /** Comments/PIs/whitespace before the root element, in order. */
  prolog: XmlNode[] = [];
  /** Comments/PIs/whitespace after the root element, in order. */
  epilog: XmlNode[] = [];
  #root: XmlElement;
  #dirty = false;

  constructor(root: XmlElement) {
    this.#root = root;
    root.docIfRoot = this;
  }

  get root(): XmlElement {
    return this.#root;
  }

  get dirty(): boolean {
    return this.#dirty;
  }

  markDirty(): void {
    this.#dirty = true;
  }
}
