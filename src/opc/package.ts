/**
 * OPC package objects — port of pptx/opc/package.py with the ts-pptx
 * preserve-first save path:
 *
 *   - every part retains its original blob; XML parses lazily on first use
 *   - untouched parts/rels/content-types are written back byte-identical
 *   - zip entries unreachable from the rels graph (python-pptx silently
 *     DROPS these) are carried through verbatim
 *   - output entry order follows the original archive, new items appended
 */
import { InvalidXmlError, XlsxError } from "../exc.js";
import { XmlDocument, XmlElement } from "../xml/dom.js";
import { parseXml } from "../xml/parser.js";
import { serializeDocument } from "../xml/serializer.js";
import { countAttrValues } from "../oxml/base.js";
import { nsmap } from "../oxml/ns.js";
import { RELATIONSHIP_TYPE as RT } from "./constants.js";
import { ContentTypeMap } from "./content-types.js";
import { CT_Relationships } from "./oxml.js";
import { CONTENT_TYPES_URI, PACKAGE_URI, PackURI } from "./packuri.js";
import { Relationships, Relationship } from "./rels.js";
import { PackageItem, PackageReader, PackageWriter } from "./serialized.js";

export class Part {
  #partname: PackURI;
  #contentType: string;
  #blob: Uint8Array;
  #blobDirty = false;
  readonly pkg: OpcPackage;
  readonly rels: Relationships;

  constructor(partname: PackURI, contentType: string, pkg: OpcPackage, blob: Uint8Array) {
    this.#partname = partname;
    this.#contentType = contentType;
    this.pkg = pkg;
    this.#blob = blob;
    this.rels = new Relationships(partname.baseURI);
  }

  static load(partname: PackURI, contentType: string, pkg: OpcPackage, blob: Uint8Array): Part {
    return new this(partname, contentType, pkg, blob);
  }

  get partname(): PackURI {
    return this.#partname;
  }

  /** Rename the part. Marks the package structure dirty. */
  setPartname(partname: PackURI): void {
    this.#partname = partname;
    this.pkg.markStructureDirty();
  }

  get contentType(): string {
    return this.#contentType;
  }

  get blob(): Uint8Array {
    return this.#blob;
  }

  setBlob(blob: Uint8Array): void {
    this.#blob = blob;
    this.#blobDirty = true;
  }

  get isDirty(): boolean {
    return this.#blobDirty;
  }

  /** Original bytes as loaded from the package (save-path internal). */
  protected get originalBlob(): Uint8Array {
    return this.#blob;
  }

  loadRelsFromXml(
    xmlRels: CT_Relationships,
    parts: ReadonlyMap<string, Part>,
    originalXml: Uint8Array | null,
  ): void {
    this.rels.loadFromXml(this.#partname.baseURI, xmlRels, parts, originalXml);
  }

  // -- _RelatableMixin -------------------------------------------------

  partRelatedBy(reltype: string): Part {
    return this.rels.partWithReltype(reltype);
  }

  relateTo(target: Part | string, reltype: string, isExternal = false): string {
    if (typeof target === "string") {
      if (!isExternal) throw new XlsxError("string target requires isExternal");
      return this.rels.getOrAddExtRel(reltype, target);
    }
    return this.rels.getOrAdd(reltype, target);
  }

  relatedPart(rId: string): Part {
    return this.rels.get(rId).targetPart;
  }

  targetRef(rId: string): string {
    return this.rels.get(rId).targetRef;
  }

  dropRel(rId: string): void {
    this.rels.pop(rId);
  }
}

export class XmlPart extends Part {
  #doc: XmlDocument | null = null;

  static override load(
    partname: PackURI,
    contentType: string,
    pkg: OpcPackage,
    blob: Uint8Array,
  ): XmlPart {
    // `new this` so registered subclasses (SlidePart, ...) construct themselves
    return new this(partname, contentType, pkg, blob);
  }

  /** Parsed XML document — parses the original blob on first access. */
  get doc(): XmlDocument {
    if (this.#doc === null) {
      try {
        this.#doc = parseXml(this.originalBlob);
      } catch (e) {
        throw new InvalidXmlError(`part ${this.partname.uri}: ${(e as Error).message}`);
      }
    }
    return this.#doc;
  }

  get root(): XmlElement {
    return this.doc.root;
  }

  override get blob(): Uint8Array {
    if (this.#doc === null || !this.#doc.dirty) return this.originalBlob;
    return serializeDocument(this.#doc);
  }

  override setBlob(_blob: Uint8Array): void {
    throw new XlsxError("XmlPart blob is derived from its XML; edit the document instead");
  }

  override get isDirty(): boolean {
    return this.#doc?.dirty ?? false;
  }

  /** Drop rel `rId` unless the part's XML references it more than once —
   * port of XmlPart.drop_rel's `//@r:id` reference count. */
  override dropRel(rId: string): void {
    if (this.relRefCount(rId) < 2) this.rels.pop(rId);
  }

  protected relRefCount(rId: string): number {
    return countAttrValues(this.root, nsmap.r, "id", rId);
  }
}

export interface PartConstructor {
  load(partname: PackURI, contentType: string, pkg: OpcPackage, blob: Uint8Array): Part;
}

export class PartFactory {
  /** content-type → Part subtype registry, populated by src/index.ts
   * (mirrors python-pptx's pptx/__init__.py registrations). */
  static partTypeFor = new Map<string, PartConstructor>();

  static create(partname: PackURI, contentType: string, pkg: OpcPackage, blob: Uint8Array): Part {
    const registered = PartFactory.partTypeFor.get(contentType);
    if (registered) return registered.load(partname, contentType, pkg, blob);
    // Unregistered XML content types get a (lazy) XmlPart so fidelity holds
    // for touched-but-unmodeled parts; everything else is a binary Part.
    if (contentType.endsWith("+xml") || contentType === "application/xml") {
      return XmlPart.load(partname, contentType, pkg, blob);
    }
    return Part.load(partname, contentType, pkg, blob);
  }
}

export class OpcPackage {
  readonly rels = new Relationships("/");
  #contentTypes!: ContentTypeMap;
  #opaqueItems = new Map<string, Uint8Array>();
  #originalEntryOrder: string[] = [];
  #structureDirty = false;

  protected constructor() {}

  static async open(source: string | Uint8Array | ArrayBuffer): Promise<OpcPackage> {
    const pkg = new this();
    await pkg.load(source);
    return pkg;
  }

  protected async load(source: string | Uint8Array | ArrayBuffer): Promise<void> {
    const reader = await PackageReader.open(source);
    this.#originalEntryOrder = [...reader.entryNames()];
    this.#contentTypes = ContentTypeMap.fromXml(reader.read(CONTENT_TYPES_URI));

    // -- depth-first rels-graph walk (port of _PackageLoader._xml_rels) --
    const xmlRels = new Map<string, CT_Relationships>(); // partname uri -> parsed rels
    const relsBytes = new Map<string, Uint8Array | null>(); // partname uri -> original rels bytes
    const visit = (sourceUri: PackURI): void => {
      const raw = reader.relsXmlFor(sourceUri);
      relsBytes.set(sourceUri.uri, raw);
      const rels =
        raw === null
          ? CT_Relationships.new()
          : CT_Relationships.fromRoot(parseXml(raw).root);
      xmlRels.set(sourceUri.uri, rels);
      for (const rel of rels.relationshipLst) {
        if (rel.targetMode === "External") continue;
        const target = PackURI.fromRelRef(sourceUri.baseURI, rel.targetRef);
        if (xmlRels.has(target.uri)) continue;
        visit(target);
      }
    };
    visit(PACKAGE_URI);

    // -- construct parts (skip partnames missing from the archive) --
    const parts = new Map<string, Part>();
    for (const uri of xmlRels.keys()) {
      if (uri === "/") continue;
      const partname = new PackURI(uri);
      if (!reader.has(partname)) continue;
      parts.set(
        uri,
        PartFactory.create(partname, this.#contentTypes.lookup(partname), this, reader.read(partname)),
      );
    }

    // -- load relationships into each part, then the package --
    for (const [uri, part] of parts) {
      part.loadRelsFromXml(xmlRels.get(uri)!, parts, relsBytes.get(uri) ?? null);
    }
    this.rels.loadFromXml("/", xmlRels.get("/")!, parts, relsBytes.get("/") ?? null);

    // -- opaque sweep: archive entries owned by nothing above --
    const owned = new Set<string>([CONTENT_TYPES_URI.membername]);
    for (const uri of xmlRels.keys()) {
      const packUri = uri === "/" ? PACKAGE_URI : new PackURI(uri);
      if (relsBytes.get(uri) !== null) owned.add(packUri.relsUri.membername);
    }
    for (const part of parts.values()) owned.add(part.partname.membername);
    for (const name of reader.entryNames()) {
      if (!owned.has(name)) this.#opaqueItems.set(name, reader.read(name));
    }
  }

  // -- traversal (ports of iter_rels/iter_parts) -------------------------

  *iterRels(): IterableIterator<Relationship> {
    const visited = new Set<Part>();
    function* walk(rels: Relationships): IterableIterator<Relationship> {
      for (const rel of rels) {
        yield rel;
        if (rel.isExternal) continue;
        const part = rel.targetPart;
        if (visited.has(part)) continue;
        visited.add(part);
        yield* walk(part.rels);
      }
    }
    yield* walk(this.rels);
  }

  *iterParts(): IterableIterator<Part> {
    const visited = new Set<Part>();
    for (const rel of this.iterRels()) {
      if (rel.isExternal) continue;
      const part = rel.targetPart;
      if (visited.has(part)) continue;
      visited.add(part);
      yield part;
    }
  }

  // -- _RelatableMixin ----------------------------------------------------

  partRelatedBy(reltype: string): Part {
    return this.rels.partWithReltype(reltype);
  }

  relateTo(target: Part | string, reltype: string, isExternal = false): string {
    if (typeof target === "string") {
      if (!isExternal) throw new XlsxError("string target requires isExternal");
      return this.rels.getOrAddExtRel(reltype, target);
    }
    return this.rels.getOrAdd(reltype, target);
  }

  relatedPart(rId: string): Part {
    return this.rels.get(rId).targetPart;
  }

  dropRel(rId: string): void {
    this.rels.pop(rId);
  }

  get mainDocumentPart(): Part {
    return this.partRelatedBy(RT.OFFICE_DOCUMENT);
  }

  /** Next available partname for `tmpl` like "/ppt/slides/slide{n}.xml".
   * First-gap allocation (deliberate divergence from python's len+1, which
   * collides once slide delete exists). */
  nextPartname(tmpl: string): PackURI {
    if (!tmpl.includes("{n}")) throw new XlsxError(`partname template missing {n}: ${tmpl}`);
    const taken = new Set<string>();
    for (const part of this.iterParts()) taken.add(part.partname.uri);
    for (let n = 1; ; n++) {
      const candidate = tmpl.replace("{n}", String(n));
      if (!taken.has(candidate)) return new PackURI(candidate);
    }
  }

  /** Call when a part is added, removed, renamed, or re-typed. */
  markStructureDirty(): void {
    this.#structureDirty = true;
  }

  get structureDirty(): boolean {
    return this.#structureDirty;
  }

  // -- save ---------------------------------------------------------------

  async toBuffer(): Promise<Uint8Array> {
    return PackageWriter.toBuffer(this.#serializedItems());
  }

  async save(path: string): Promise<void> {
    await PackageWriter.write(path, this.#serializedItems());
  }

  #serializedItems(): PackageItem[] {
    const parts = [...this.iterParts()];
    const items = new Map<string, Uint8Array>();

    // content-types: original bytes unless the structure changed
    const partsAndOpaque = new Set<string>(this.#opaqueItems.keys());
    for (const p of parts) partsAndOpaque.add(p.partname.membername);
    const ctBlob =
      !this.#structureDirty && this.#contentTypes.originalBlob !== null
        ? this.#contentTypes.originalBlob
        : this.#contentTypes.regeneratedBytes(parts, partsAndOpaque);
    items.set(CONTENT_TYPES_URI.membername, ctBlob);

    // package rels (always present)
    const pkgRels = this.rels.toXmlBytes();
    if (pkgRels === null) throw new XlsxError("package has no relationships");
    items.set(PACKAGE_URI.relsUri.membername, pkgRels);

    // parts + their rels items
    for (const part of parts) {
      items.set(part.partname.membername, part.blob);
      const relsXml = part.rels.toXmlBytes();
      if (relsXml !== null) items.set(part.partname.relsUri.membername, relsXml);
    }

    // opaque passthrough (never dropped, unlike python-pptx)
    for (const [name, blob] of this.#opaqueItems) {
      if (!items.has(name)) items.set(name, blob);
    }

    // original entry order first, then any new items in insertion order
    const ordered: PackageItem[] = [];
    const emitted = new Set<string>();
    for (const name of this.#originalEntryOrder) {
      const blob = items.get(name);
      if (blob === undefined) continue;
      ordered.push({ membername: name, blob });
      emitted.add(name);
    }
    for (const [name, blob] of items) {
      if (!emitted.has(name)) ordered.push({ membername: name, blob });
    }
    return ordered;
  }
}
