/**
 * PackURI — port of pptx/opc/packuri.py.
 *
 * python subclasses str; TS uses an immutable value class. All Map keys in
 * ts-pptx use the plain string form (`.uri`).
 */
import path from "node:path";

const FILENAME_RE = /^([a-zA-Z]+)([0-9][0-9]*)?/;

export class PackURI {
  readonly uri: string;

  constructor(uri: string) {
    if (!uri.startsWith("/")) {
      throw new Error(`PackURI must begin with slash, got ${JSON.stringify(uri)}`);
    }
    this.uri = uri;
  }

  /** Absolute pack URI formed by translating `relativeRef` onto `baseURI`.
   * A `relativeRef` that is already package-absolute (leading "/") is taken
   * as-is — Excel authors rels with absolute targets ("/xl/worksheets/sheet1.xml")
   * where PowerPoint/Word use relative ones; OPC permits both. */
  static fromRelRef(baseURI: string, relativeRef: string): PackURI {
    if (relativeRef.startsWith("/")) return new PackURI(path.posix.normalize(relativeRef));
    const joined = path.posix.join(baseURI, relativeRef);
    return new PackURI(path.posix.normalize(joined));
  }

  /** Directory portion: "/ppt/slides" for "/ppt/slides/slide1.xml"; "/" for "/". */
  get baseURI(): string {
    return path.posix.dirname(this.uri);
  }

  /** Extension without the period: "xml" for "/ppt/slides/slide1.xml". */
  get ext(): string {
    const raw = path.posix.extname(this.uri);
    return raw.startsWith(".") ? raw.slice(1) : raw;
  }

  /** "slide1.xml" for "/ppt/slides/slide1.xml"; "" for "/". */
  get filename(): string {
    return this.uri === "/" ? "" : path.posix.basename(this.uri);
  }

  /** Integer partname index: 21 for ".../slide21.xml", null for singletons. */
  get idx(): number | null {
    const filename = this.filename;
    if (!filename) return null;
    const stem = filename.replace(/\.[^.]*$/, "");
    const m = FILENAME_RE.exec(stem);
    if (m === null || !m[2]) return null;
    return parseInt(m[2], 10);
  }

  /** Zip member name: the pack URI without the leading slash. */
  get membername(): string {
    return this.uri.slice(1);
  }

  /** Relative reference to this item from `baseURI`, e.g. "../media/image1.png". */
  relativeRef(baseURI: string): string {
    return baseURI === "/" ? this.uri.slice(1) : path.posix.relative(baseURI, this.uri);
  }

  /** Partname of the corresponding .rels item. */
  get relsUri(): PackURI {
    const base = this.baseURI;
    return new PackURI(path.posix.join(base, "_rels", `${this.filename}.rels`));
  }

  toString(): string {
    return this.uri;
  }

  equals(other: PackURI | string): boolean {
    return this.uri === (typeof other === "string" ? other : other.uri);
  }
}

export const PACKAGE_URI = new PackURI("/");
export const CONTENT_TYPES_URI = new PackURI("/[Content_Types].xml");
