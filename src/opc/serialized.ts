/**
 * Physical package read/write over jszip — port of pptx/opc/serialized.py
 * (_ZipPkgReader/_ZipPkgWriter; directory packages deferred).
 */
import { readFile, writeFile } from "node:fs/promises";
import JSZip from "jszip";
import { KeyLookupError, PackageNotFoundError } from "../exc.js";
import { PackURI } from "./packuri.js";

/** Fixed timestamp for deterministic zip output run-to-run (DOS zip time
 * cannot represent the 1970 epoch, so a fixed post-1980 date is used). */
const ZIP_DATE = new Date(2000, 0, 1);

export class PackageReader {
  #entries: Map<string, Uint8Array>; // keyed by zip membername
  #order: string[]; // original entry order

  private constructor(entries: Map<string, Uint8Array>, order: string[]) {
    this.#entries = entries;
    this.#order = order;
  }

  static async open(source: string | Uint8Array | ArrayBuffer): Promise<PackageReader> {
    let bytes: Uint8Array;
    if (typeof source === "string") {
      try {
        bytes = new Uint8Array(await readFile(source));
      } catch {
        throw new PackageNotFoundError(`Package not found at '${source}'`);
      }
    } else if (source instanceof ArrayBuffer) {
      bytes = new Uint8Array(source);
    } else {
      bytes = source;
    }
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(bytes);
    } catch (e) {
      throw new PackageNotFoundError(`source is not a zip package: ${(e as Error).message}`);
    }
    const entries = new Map<string, Uint8Array>();
    const order: string[] = [];
    for (const name of Object.keys(zip.files)) {
      const f = zip.files[name]!;
      if (f.dir) continue;
      entries.set(name, await f.async("uint8array"));
      order.push(name);
    }
    return new PackageReader(entries, order);
  }

  has(uri: PackURI | string): boolean {
    return this.#entries.has(membernameOf(uri));
  }

  read(uri: PackURI | string): Uint8Array {
    const blob = this.#entries.get(membernameOf(uri));
    if (blob === undefined) throw new KeyLookupError(`no member '${membernameOf(uri)}' in package`);
    return blob;
  }

  /** Rels-item bytes for `partname`, or null when the part has no rels item. */
  relsXmlFor(partname: PackURI): Uint8Array | null {
    const name = partname.relsUri.membername;
    return this.#entries.get(name) ?? null;
  }

  /** Zip membernames in original archive order. */
  entryNames(): readonly string[] {
    return this.#order;
  }
}

function membernameOf(uri: PackURI | string): string {
  return typeof uri === "string" ? uri.replace(/^\//, "") : uri.membername;
}

export interface PackageItem {
  membername: string;
  blob: Uint8Array;
}

export class PackageWriter {
  static async toBuffer(items: Iterable<PackageItem>): Promise<Uint8Array> {
    const zip = new JSZip();
    for (const { membername, blob } of items) {
      zip.file(membername, blob, { binary: true, date: ZIP_DATE, createFolders: false });
    }
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  static async write(path: string, items: Iterable<PackageItem>): Promise<void> {
    await writeFile(path, await PackageWriter.toBuffer(items));
  }
}
