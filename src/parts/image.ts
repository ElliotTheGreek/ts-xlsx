/** Image part (xl/media/imageN.ext) — binary, sha1-deduplicated. */
import { createHash } from "node:crypto";
import { OpcPackage, Part } from "../opc/package.js";
import { PackURI } from "../opc/packuri.js";
import { probeImage } from "../image/probe.js";

function sha1Of(blob: Uint8Array): string {
  return createHash("sha1").update(blob).digest("hex");
}

export class ImagePart extends Part {
  #sha1: string;

  constructor(partname: PackURI, contentType: string, pkg: OpcPackage, blob: Uint8Array) {
    super(partname, contentType, pkg, blob);
    this.#sha1 = sha1Of(blob);
  }

  get sha1(): string {
    return this.#sha1;
  }

  override setBlob(blob: Uint8Array): void {
    super.setBlob(blob);
    this.#sha1 = sha1Of(blob);
  }

  static createFromBlob(pkg: OpcPackage, blob: Uint8Array): ImagePart {
    const probe = probeImage(blob);
    const partname = pkg.nextPartname(`/xl/media/image{n}.${probe.ext}`);
    return new ImagePart(partname, probe.contentType, pkg, blob);
  }
}

/** Reuse an identical (sha1-equal) image already embedded, else create one. */
export function getOrAddImagePart(pkg: OpcPackage, blob: Uint8Array): ImagePart {
  const sha1 = sha1Of(blob);
  for (const part of pkg.iterParts()) {
    if (part instanceof ImagePart && part.sha1 === sha1) return part;
  }
  const part = ImagePart.createFromBlob(pkg, blob);
  pkg.markStructureDirty();
  return part;
}
