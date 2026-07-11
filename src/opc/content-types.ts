/**
 * Content-types map — merges python-pptx's _ContentTypeMap (read side,
 * package.py) and _ContentTypesItem (write side, serialized.py), plus the
 * ts-pptx preservation rule: the original [Content_Types].xml bytes are
 * written back verbatim unless the package structure changed.
 */
import { KeyLookupError } from "../exc.js";
import { parseXml } from "../xml/parser.js";
import { CONTENT_TYPE as CT } from "./constants.js";
import { CT_Types } from "./oxml.js";
import { PackURI } from "./packuri.js";
import { defaultContentTypes } from "./spec.js";

interface OverrideEntry {
  partName: string; // original case, for re-emission
  contentType: string;
}

interface DefaultEntry {
  extension: string; // original case, for re-emission
  contentType: string;
}

export interface ContentTypedPart {
  readonly partname: PackURI;
  readonly contentType: string;
}

export class ContentTypeMap {
  #overrides: Map<string, OverrideEntry>; // keyed by lowercased partname
  #defaults: Map<string, DefaultEntry>; // keyed by lowercased extension
  #originalBlob: Uint8Array | null;

  private constructor(
    overrides: Map<string, OverrideEntry>,
    defaults: Map<string, DefaultEntry>,
    originalBlob: Uint8Array | null,
  ) {
    this.#overrides = overrides;
    this.#defaults = defaults;
    this.#originalBlob = originalBlob;
  }

  static fromXml(blob: Uint8Array): ContentTypeMap {
    const types = CT_Types.fromRoot(parseXml(blob).root);
    const overrides = new Map<string, OverrideEntry>();
    for (const o of types.overrideLst) {
      overrides.set(o.partName.toLowerCase(), { partName: o.partName, contentType: o.contentType });
    }
    const defaults = new Map<string, DefaultEntry>();
    for (const d of types.defaultLst) {
      defaults.set(d.extension.toLowerCase(), {
        extension: d.extension,
        contentType: d.contentType,
      });
    }
    return new ContentTypeMap(overrides, defaults, blob);
  }

  /** Content type for `partname`: override first, then extension default. */
  lookup(partname: PackURI): string {
    const override = this.#overrides.get(partname.uri.toLowerCase());
    if (override) return override.contentType;
    const dflt = this.#defaults.get(partname.ext.toLowerCase());
    if (dflt) return dflt.contentType;
    throw new KeyLookupError(
      `no content-type for partname '${partname.uri}' in [Content_Types].xml`,
    );
  }

  /** Original bytes, written verbatim when the package structure is unchanged. */
  get originalBlob(): Uint8Array | null {
    return this.#originalBlob;
  }

  /**
   * Regenerate [Content_Types].xml. python-pptx rebuilds purely from the
   * rels-reachable parts; ts-pptx additionally (a) starts from the original
   * entries so orphan/opaque package items keep their content types, and
   * (b) drops overrides whose partname no longer exists.
   *
   * `existingMembernames` = every membername that will be written (parts +
   * opaque items), used to filter stale overrides.
   */
  regeneratedBytes(
    parts: readonly ContentTypedPart[],
    existingMembernames: ReadonlySet<string>,
  ): Uint8Array {
    const defaults = new Map<string, DefaultEntry>();
    defaults.set("rels", { extension: "rels", contentType: CT.OPC_RELATIONSHIPS });
    defaults.set("xml", { extension: "xml", contentType: CT.XML });
    for (const [k, v] of this.#defaults) defaults.set(k, v);

    const overrides = new Map<string, OverrideEntry>();
    for (const [k, v] of this.#overrides) {
      if (existingMembernames.has(v.partName.replace(/^\//, ""))) overrides.set(k, v);
    }

    for (const part of parts) {
      const ext = part.partname.ext;
      const isDefault = defaultContentTypes.some(
        ([e, ct]) => e === ext.toLowerCase() && ct === part.contentType,
      );
      if (isDefault) {
        defaults.set(ext.toLowerCase(), { extension: ext, contentType: part.contentType });
      } else {
        overrides.set(part.partname.uri.toLowerCase(), {
          partName: part.partname.uri,
          contentType: part.contentType,
        });
      }
    }

    const types = CT_Types.new();
    for (const key of [...defaults.keys()].sort()) {
      const d = defaults.get(key)!;
      types.addDefault(d.extension, d.contentType);
    }
    for (const key of [...overrides.keys()].sort()) {
      const o = overrides.get(key)!;
      types.addOverride(new PackURI(o.partName), o.contentType);
    }
    return types.xmlFileBytes();
  }
}
