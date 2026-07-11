/**
 * Relationships — port of _Relationships/_Relationship (pptx/opc/package.py)
 * with the ts-pptx preservation rule: original .rels bytes are written back
 * verbatim unless a relationship was added or removed, in which case the
 * item is regenerated in python-pptx's deterministic numeric-rId order.
 */
import { KeyLookupError, XlsxError } from "../exc.js";
import { RELATIONSHIP_TARGET_MODE as RTM } from "./constants.js";
import { CT_Relationships } from "./oxml.js";
import { PackURI } from "./packuri.js";
import type { Part } from "./package.js";

export class Relationship {
  constructor(
    private readonly baseUri: string,
    readonly rId: string,
    readonly reltype: string,
    private readonly targetMode: "Internal" | "External",
    private readonly target: Part | string,
  ) {}

  get isExternal(): boolean {
    return this.targetMode === RTM.EXTERNAL;
  }

  get targetPart(): Part {
    if (this.isExternal) {
      throw new XlsxError("`.targetPart` is undefined when target-mode is external");
    }
    return this.target as Part;
  }

  get targetPartname(): PackURI {
    return this.targetPart.partname;
  }

  /** Relative partname for internal rels; the URL for external rels. */
  get targetRef(): string {
    if (this.isExternal) return this.target as string;
    return this.targetPartname.relativeRef(this.baseUri);
  }
}

export class Relationships implements Iterable<Relationship> {
  #rels = new Map<string, Relationship>();
  #baseUri: string;
  #originalXml: Uint8Array | null = null;
  #dirty = false;

  constructor(baseUri: string) {
    this.#baseUri = baseUri;
  }

  /**
   * Replace contents from parsed rels XML. Internal rels whose target
   * partname is absent from `parts` are skipped ("voided" targets like
   * /ppt/slides/NULL, per python-pptx) — but the original bytes are still
   * written back verbatim as long as this collection stays untouched.
   */
  loadFromXml(
    baseUri: string,
    xmlRels: CT_Relationships,
    parts: ReadonlyMap<string, Part>,
    originalXml: Uint8Array | null,
  ): void {
    this.#rels.clear();
    this.#baseUri = baseUri;
    this.#originalXml = originalXml;
    this.#dirty = false;
    for (const relElm of xmlRels.relationshipLst) {
      if (relElm.targetMode === RTM.INTERNAL) {
        const partname = PackURI.fromRelRef(baseUri, relElm.targetRef);
        const part = parts.get(partname.uri);
        if (part === undefined) continue;
        this.#rels.set(
          relElm.rId,
          new Relationship(baseUri, relElm.rId, relElm.reltype, RTM.INTERNAL, part),
        );
      } else {
        this.#rels.set(
          relElm.rId,
          new Relationship(baseUri, relElm.rId, relElm.reltype, RTM.EXTERNAL, relElm.targetRef),
        );
      }
    }
  }

  get(rId: string): Relationship {
    const rel = this.#rels.get(rId);
    if (rel === undefined) throw new KeyLookupError(`no relationship with key '${rId}'`);
    return rel;
  }

  has(rId: string): boolean {
    return this.#rels.has(rId);
  }

  get size(): number {
    return this.#rels.size;
  }

  [Symbol.iterator](): Iterator<Relationship> {
    return this.#rels.values()[Symbol.iterator]();
  }

  /** rId of a `reltype` rel to `targetPart`, adding one if not present. */
  getOrAdd(reltype: string, targetPart: Part): string {
    return this.#getMatching(reltype, targetPart, false) ?? this.#add(reltype, targetPart, false);
  }

  /** rId of an external `reltype` rel to `targetRef`, adding one if not present. */
  getOrAddExtRel(reltype: string, targetRef: string): string {
    return this.#getMatching(reltype, targetRef, true) ?? this.#add(reltype, targetRef, true);
  }

  /** Add a relationship with a caller-specified rId — used when cloning a
   * part whose XML embeds r:id references that must keep their values
   * (slide duplication). Throws if the rId is taken. */
  addWithRId(rId: string, reltype: string, target: Part | string, isExternal: boolean): void {
    if (this.#rels.has(rId)) throw new XlsxError(`rId '${rId}' already present`);
    this.#rels.set(
      rId,
      new Relationship(
        this.#baseUri,
        rId,
        reltype,
        isExternal ? RTM.EXTERNAL : RTM.INTERNAL,
        target,
      ),
    );
    this.#dirty = true;
  }

  /** Re-point an existing internal relationship at `target`, keeping its
   * rId and reltype (used by Picture.replaceImage — the referencing XML
   * stays untouched). Marks the collection dirty. */
  replaceTarget(rId: string, target: Part): void {
    const old = this.get(rId);
    if (old.isExternal) throw new XlsxError("cannot retarget an external relationship");
    this.#rels.set(
      rId,
      new Relationship(this.#baseUri, rId, old.reltype, RTM.INTERNAL, target),
    );
    this.#dirty = true;
  }

  pop(rId: string): Relationship {
    const rel = this.get(rId);
    this.#rels.delete(rId);
    this.#dirty = true;
    return rel;
  }

  /** Target part of the single rel of `reltype`; KeyLookupError if none,
   * XlsxError if more than one (python: KeyError/ValueError). */
  partWithReltype(reltype: string): Part {
    const matches = [...this.#rels.values()].filter((r) => r.reltype === reltype);
    if (matches.length === 0) {
      throw new KeyLookupError(`no relationship of type '${reltype}' in collection`);
    }
    if (matches.length > 1) {
      throw new XlsxError(`multiple relationships of type '${reltype}' in collection`);
    }
    return matches[0]!.targetPart;
  }

  get dirty(): boolean {
    return this.#dirty;
  }

  /** Bytes for the .rels item, or null when no item should be written. */
  toXmlBytes(): Uint8Array | null {
    if (!this.#dirty && this.#originalXml !== null) return this.#originalXml;
    if (this.#rels.size === 0) return null;
    const relsElm = CT_Relationships.new();
    for (const rel of this.#inNumericalOrder()) {
      relsElm.addRel(rel.rId, rel.reltype, rel.targetRef, rel.isExternal);
    }
    return relsElm.xmlFileBytes();
  }

  #inNumericalOrder(): Relationship[] {
    // python: sort by (numeric rId suffix or 0, rId string)
    return [...this.#rels.values()].sort((a, b) => {
      const na = numOf(a.rId);
      const nb = numOf(b.rId);
      if (na !== nb) return na - nb;
      return a.rId < b.rId ? -1 : a.rId > b.rId ? 1 : 0;
    });
  }

  #getMatching(reltype: string, target: Part | string, isExternal: boolean): string | null {
    for (const rel of this.#rels.values()) {
      if (rel.reltype !== reltype || rel.isExternal !== isExternal) continue;
      const relTarget = rel.isExternal ? rel.targetRef : rel.targetPart;
      if (relTarget === target) return rel.rId;
    }
    return null;
  }

  #add(reltype: string, target: Part | string, isExternal: boolean): string {
    const rId = this.#nextRId();
    this.#rels.set(
      rId,
      new Relationship(
        this.#baseUri,
        rId,
        reltype,
        isExternal ? RTM.EXTERNAL : RTM.INTERNAL,
        target,
      ),
    );
    this.#dirty = true;
    return rId;
  }

  /** First unused "rIdN", filling gaps — python's countdown search, exactly. */
  #nextRId(): string {
    for (let n = this.#rels.size + 1; n > 0; n--) {
      const candidate = `rId${n}`;
      if (!this.#rels.has(candidate)) return candidate;
    }
    throw new XlsxError("impossible: more distinct rIds than relationships");
  }
}

function numOf(rId: string): number {
  return rId.startsWith("rId") && /^[0-9]+$/.test(rId.slice(3)) ? parseInt(rId.slice(3), 10) : 0;
}
