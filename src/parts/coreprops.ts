/** Core document properties — port of docx/opc/coreprops.py (read + write). */
import { XmlPart } from "../opc/package.js";
import { XmlElement } from "../xml/dom.js";
import { createElement } from "../oxml/base.js";
import { nsmap } from "../oxml/ns.js";

export class CorePropertiesPart extends XmlPart {
  get coreProperties(): CoreProperties {
    return new CoreProperties(this.root);
  }
}

/** W3CDTF (ISO-8601, seconds, trailing Z) for a Date. */
function toW3CDTF(d: Date): string {
  return d.toISOString().replace(/\.\d{3}Z$/, "Z");
}

export class CoreProperties {
  constructor(readonly element: XmlElement) {}

  // -- reads -------------------------------------------------------------
  #text(nsUri: string, localName: string): string {
    return this.element.find(nsUri, localName)?.text ?? "";
  }
  #date(localName: string): Date | undefined {
    const raw = this.#text(nsmap.dcterms, localName);
    if (raw === "") return undefined;
    const d = new Date(raw);
    return Number.isNaN(d.getTime()) ? undefined : d;
  }

  // -- writes ------------------------------------------------------------
  #getOrCreate(tag: `${"dc" | "cp" | "dcterms"}:${string}`): XmlElement {
    const { nsUri, localName } = split(tag);
    let el = this.element.find(nsUri, localName);
    if (!el) {
      el = createElement(tag, this.element);
      this.element.appendChild(el);
    }
    return el;
  }
  #setText(tag: `${"dc" | "cp"}:${string}`, value: string): void {
    if (value === "") {
      const { nsUri, localName } = split(tag);
      const el = this.element.find(nsUri, localName);
      if (el) this.element.removeChild(el);
      return;
    }
    this.#getOrCreate(tag).setText(value);
  }
  #setDate(localName: `dcterms:${string}`, value: Date | undefined): void {
    if (value === undefined) {
      const { nsUri, localName: ln } = split(localName);
      const el = this.element.find(nsUri, ln);
      if (el) this.element.removeChild(el);
      return;
    }
    const el = this.#getOrCreate(localName);
    el.setAttr("xsi:type", "dcterms:W3CDTF");
    el.setText(toW3CDTF(value));
  }

  get title(): string {
    return this.#text(nsmap.dc, "title");
  }
  set title(v: string) {
    this.#setText("dc:title", v);
  }
  get subject(): string {
    return this.#text(nsmap.dc, "subject");
  }
  set subject(v: string) {
    this.#setText("dc:subject", v);
  }
  get author(): string {
    return this.#text(nsmap.dc, "creator");
  }
  set author(v: string) {
    this.#setText("dc:creator", v);
  }
  get comments(): string {
    return this.#text(nsmap.dc, "description");
  }
  set comments(v: string) {
    this.#setText("dc:description", v);
  }
  get keywords(): string {
    return this.#text(nsmap.cp, "keywords");
  }
  set keywords(v: string) {
    this.#setText("cp:keywords", v);
  }
  get category(): string {
    return this.#text(nsmap.cp, "category");
  }
  set category(v: string) {
    this.#setText("cp:category", v);
  }
  get lastModifiedBy(): string {
    return this.#text(nsmap.cp, "lastModifiedBy");
  }
  set lastModifiedBy(v: string) {
    this.#setText("cp:lastModifiedBy", v);
  }
  get contentStatus(): string {
    return this.#text(nsmap.cp, "contentStatus");
  }
  set contentStatus(v: string) {
    this.#setText("cp:contentStatus", v);
  }

  /** 0 when absent or non-numeric (python parity). */
  get revision(): number {
    const raw = this.#text(nsmap.cp, "revision");
    const v = parseInt(raw, 10);
    return Number.isNaN(v) ? 0 : v;
  }
  set revision(value: number) {
    this.#getOrCreate("cp:revision").setText(String(value));
  }

  get created(): Date | undefined {
    return this.#date("created");
  }
  set created(v: Date | undefined) {
    this.#setDate("dcterms:created", v);
  }
  get modified(): Date | undefined {
    return this.#date("modified");
  }
  set modified(v: Date | undefined) {
    this.#setDate("dcterms:modified", v);
  }
  get lastPrinted(): Date | undefined {
    return this.#date("lastPrinted");
  }
}

function split(tag: string): { nsUri: string; localName: string } {
  const c = tag.indexOf(":");
  const prefix = tag.slice(0, c) as keyof typeof nsmap;
  return { nsUri: nsmap[prefix], localName: tag.slice(c + 1) };
}
