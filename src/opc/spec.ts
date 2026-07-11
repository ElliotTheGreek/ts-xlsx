/** SpreadsheetML default content-type mappings (ECMA-376 / openpyxl). */
import { CONTENT_TYPE as CT } from "./constants.js";

/** (extension, content type) pairs eligible for a <Default> entry when
 * [Content_Types].xml is regenerated. Excel emits `rels` and `xml` plus a
 * default per image/binary extension actually present in the package. */
export const defaultContentTypes: ReadonlyArray<readonly [string, string]> = [
  ["bin", CT.SML_PRINTER_SETTINGS],
  ["bmp", CT.BMP],
  ["emf", CT.X_EMF],
  ["gif", CT.GIF],
  ["jpe", CT.JPEG],
  ["jpeg", CT.JPEG],
  ["jpg", CT.JPEG],
  ["png", CT.PNG],
  ["rels", CT.OPC_RELATIONSHIPS],
  ["tif", CT.TIFF],
  ["tiff", CT.TIFF],
  ["vml", CT.OFC_VML_DRAWING],
  ["wdp", CT.MS_PHOTO],
  ["wmf", CT.X_WMF],
  ["xml", CT.XML],
];

export const imageContentTypes: Readonly<Record<string, string>> = {
  bmp: CT.BMP,
  emf: CT.X_EMF,
  gif: CT.GIF,
  jpe: CT.JPEG,
  jpeg: CT.JPEG,
  jpg: CT.JPEG,
  png: CT.PNG,
  tif: CT.TIFF,
  tiff: CT.TIFF,
  wdp: CT.MS_PHOTO,
  wmf: CT.X_WMF,
};
