/**
 * Image header probe — pixel dimensions + dpi for PNG, JPEG, GIF, BMP, and
 * TIFF (both byte orders).
 *
 * This python-pptx version delegates image inspection to Pillow
 * (pptx/parts/image.py Image._pil_props), so the format specs are ported
 * here, not python code. The dpi post-processing is an exact port of
 * Image.dpi's int_dpi/normalize_pil_dpi (parts/image.py:186-217): a missing
 * or non-numeric dpi yields 72, values round to int, and anything outside
 * 1..2048 falls back to 72.
 */
import { XlsxError } from "../exc.js";
import { CONTENT_TYPE as CT } from "../opc/constants.js";

export interface ImageProbe {
  /** MIME type, e.g. "image/png" (Image.content_type). */
  contentType: string;
  /** Canonical lowercase extension for the detected format (Image.ext). */
  ext: string;
  pxWidth: number;
  pxHeight: number;
  horzDpi: number;
  vertDpi: number;
}

/** Sniff `blob`'s format from its header and read size + dpi.
 * Throws XlsxError for unrecognized bytes (python Image.ext ValueError). */
export function probeImage(blob: Uint8Array): ImageProbe {
  if (startsWith(blob, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return probePng(blob);
  if (blob.length >= 2 && blob[0] === 0xff && blob[1] === 0xd8) return probeJpeg(blob);
  if (ascii(blob, 0, 6) === "GIF87a" || ascii(blob, 0, 6) === "GIF89a") return probeGif(blob);
  if (startsWith(blob, [0x49, 0x49, 0x2a, 0x00]) || startsWith(blob, [0x4d, 0x4d, 0x00, 0x2a])) {
    return probeTiff(blob);
  }
  if (ascii(blob, 0, 2) === "BM") return probeBmp(blob);
  throw new XlsxError(
    "unsupported image format, expected one of: PNG, JPEG, GIF, BMP, TIFF",
  );
}

// ---------------------------------------------------------------------------
// per-format readers
// ---------------------------------------------------------------------------

/** PNG: IHDR width/height; pHYs pixels-per-unit with unit 1 (meter) → dpi =
 * ppm * 0.0254 (Pillow PngImagePlugin chunk_pHYs; unit 0 carries no dpi). */
function probePng(b: Uint8Array): ImageProbe {
  const dv = view(b);
  let w = -1;
  let h = -1;
  let hDpi: number | undefined;
  let vDpi: number | undefined;
  let pos = 8;
  while (pos + 8 <= b.length) {
    const len = dv.getUint32(pos);
    const type = ascii(b, pos + 4, 4);
    const data = pos + 8;
    if (data + len > b.length) break;
    if (type === "IHDR" && len >= 8) {
      w = dv.getUint32(data);
      h = dv.getUint32(data + 4);
    } else if (type === "pHYs" && len >= 9) {
      const ppmX = dv.getUint32(data);
      const ppmY = dv.getUint32(data + 4);
      if (b[data + 8] === 1) {
        // unit 1 = pixels per meter
        hDpi = ppmX * 0.0254;
        vDpi = ppmY * 0.0254;
      }
    } else if (type === "IDAT" || type === "IEND") {
      break; // pHYs must precede IDAT (PNG spec §5.6)
    }
    pos = data + len + 4; // skip data + CRC
  }
  if (w < 0) throw new XlsxError("PNG image has no IHDR chunk");
  return result(CT.PNG, "png", w, h, hDpi, vDpi);
}

/** Markers whose segment is a frame header carrying height/width: SOF0-3,
 * SOF5-7, SOF9-11, SOF13-15 (C4 = DHT, C8 = JPG extension, CC = DAC). */
const SOF_MARKERS = new Set([
  0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf,
]);

/** JPEG: scan the marker stream for a SOF segment (precision byte, then
 * height/width as 16-bit BE) and the JFIF APP0 density (unit 1 = dpi,
 * unit 2 = dots/cm → × 2.54). */
function probeJpeg(b: Uint8Array): ImageProbe {
  const dv = view(b);
  let unit: number | undefined;
  let xDensity: number | undefined;
  let yDensity: number | undefined;
  let pos = 2;
  while (pos + 4 <= b.length) {
    if (b[pos] !== 0xff) throw new XlsxError("malformed JPEG marker stream");
    pos++;
    while (b[pos] === 0xff) pos++; // fill bytes
    const marker = b[pos]!;
    pos++;
    // standalone markers: TEM, RSTn, SOI (no length field)
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd8)) continue;
    if (marker === 0xd9 || marker === 0xda) break; // EOI / SOS: no SOF seen
    if (pos + 2 > b.length) break;
    const len = dv.getUint16(pos); // includes the two length bytes
    const data = pos + 2;
    if (marker === 0xe0 && len >= 14 && ascii(b, data, 5) === "JFIF\u0000") {
      unit = b[data + 7];
      xDensity = dv.getUint16(data + 8);
      yDensity = dv.getUint16(data + 10);
    }
    if (SOF_MARKERS.has(marker) && len >= 7) {
      const h = dv.getUint16(data + 1);
      const w = dv.getUint16(data + 3);
      let hDpi: number | undefined;
      let vDpi: number | undefined;
      if (unit === 1) {
        hDpi = xDensity;
        vDpi = yDensity;
      } else if (unit === 2 && xDensity !== undefined && yDensity !== undefined) {
        hDpi = xDensity * 2.54;
        vDpi = yDensity * 2.54;
      }
      return result(CT.JPEG, "jpg", w, h, hDpi, vDpi);
    }
    pos += len;
  }
  throw new XlsxError("JPEG image has no SOF frame header");
}

/** GIF: logical screen descriptor width/height (16-bit LE at offsets 6/8).
 * The format carries no density; dpi defaults to 72. */
function probeGif(b: Uint8Array): ImageProbe {
  if (b.length < 10) throw new XlsxError("GIF image is truncated");
  const w = b[6]! | (b[7]! << 8);
  const h = b[8]! | (b[9]! << 8);
  return result(CT.GIF, "gif", w, h, undefined, undefined);
}

/** BMP: BITMAPINFOHEADER width/height (int32 LE at 18/22; negative height =
 * top-down DIB) + X/YPelsPerMeter (int32 LE at 38/42) → dpi = ppm * 0.0254. */
function probeBmp(b: Uint8Array): ImageProbe {
  if (b.length < 54) throw new XlsxError("BMP image is truncated");
  const dv = view(b);
  const headerSize = dv.getUint32(14, true);
  if (headerSize < 40) throw new XlsxError("unsupported BMP header (BITMAPCOREHEADER)");
  const w = dv.getInt32(18, true);
  const h = Math.abs(dv.getInt32(22, true));
  const ppmX = dv.getInt32(38, true);
  const ppmY = dv.getInt32(42, true);
  return result(
    CT.BMP,
    "bmp",
    w,
    h,
    ppmX > 0 ? ppmX * 0.0254 : undefined,
    ppmY > 0 ? ppmY * 0.0254 : undefined,
  );
}

/** TIFF: first IFD, both byte orders. Tags: 256 ImageWidth / 257 ImageLength
 * (SHORT or LONG), 282/283 X/YResolution (RATIONAL), 296 ResolutionUnit.
 * Unit semantics mirror Pillow TiffImagePlugin: 2 (inch) or tag absent →
 * resolution is dpi; 3 (cm) → × 2.54; 1 (no absolute unit) → no dpi. */
function probeTiff(b: Uint8Array): ImageProbe {
  const le = b[0] === 0x49;
  const dv = view(b);
  const u16 = (o: number): number => dv.getUint16(o, le);
  const u32 = (o: number): number => dv.getUint32(o, le);
  const ifd = u32(4);
  if (ifd + 2 > b.length) throw new XlsxError("TIFF image is truncated");
  const count = u16(ifd);
  let w: number | undefined;
  let h: number | undefined;
  let xRes: number | undefined;
  let yRes: number | undefined;
  let unit: number | undefined;
  for (let i = 0; i < count; i++) {
    const entry = ifd + 2 + i * 12;
    if (entry + 12 > b.length) break;
    const tag = u16(entry);
    const type = u16(entry + 2);
    const valOff = entry + 8; // value is left-justified in the 4-byte field
    // SHORT (3) reads 16 bits, LONG (4) reads 32
    const intVal = (): number => (type === 3 ? u16(valOff) : u32(valOff));
    const rational = (): number | undefined => {
      if (type !== 5) return undefined;
      const o = u32(valOff);
      if (o + 8 > b.length) return undefined;
      const den = u32(o + 4);
      return den === 0 ? undefined : u32(o) / den;
    };
    if (tag === 256) w = intVal();
    else if (tag === 257) h = intVal();
    else if (tag === 282) xRes = rational();
    else if (tag === 283) yRes = rational();
    else if (tag === 296) unit = intVal();
  }
  if (w === undefined || h === undefined) {
    throw new XlsxError("TIFF image has no ImageWidth/ImageLength tags");
  }
  const toDpi = (res: number | undefined): number | undefined => {
    if (res === undefined) return undefined;
    if (unit === 3) return res * 2.54;
    if (unit === undefined || unit === 2) return res;
    return undefined; // unit 1: no absolute unit of measurement
  };
  return result(CT.TIFF, "tiff", w, h, toDpi(xRes), toDpi(yRes));
}

// ---------------------------------------------------------------------------
// shared helpers
// ---------------------------------------------------------------------------

/** Port of Image.dpi's int_dpi (parts/image.py:193-205): round to int; when
 * missing, non-numeric, < 1, or > 2048, the value is 72. */
function intDpi(dpi: number | undefined): number {
  if (dpi === undefined || !Number.isFinite(dpi)) return 72;
  const v = Math.round(dpi);
  return v < 1 || v > 2048 ? 72 : v;
}

function result(
  contentType: string,
  ext: string,
  pxWidth: number,
  pxHeight: number,
  hDpi: number | undefined,
  vDpi: number | undefined,
): ImageProbe {
  return { contentType, ext, pxWidth, pxHeight, horzDpi: intDpi(hDpi), vertDpi: intDpi(vDpi) };
}

function view(b: Uint8Array): DataView {
  return new DataView(b.buffer, b.byteOffset, b.byteLength);
}

function startsWith(b: Uint8Array, sig: readonly number[]): boolean {
  if (b.length < sig.length) return false;
  for (let i = 0; i < sig.length; i++) if (b[i] !== sig[i]) return false;
  return true;
}

function ascii(b: Uint8Array, start: number, len: number): string {
  if (start + len > b.length) return "";
  let out = "";
  for (let i = start; i < start + len; i++) out += String.fromCharCode(b[i]!);
  return out;
}
