// -- M0: fidelity backbone (lossless XML + OPC package) ------------------
export * from "./exc.js";
export * from "./xml/index.js";
export { nsmap, nsdecls, splitNsTag, nsUriOf } from "./oxml/ns.js";
export type { NsTag, NsPrefix } from "./oxml/ns.js";
export * from "./oxml/simpletypes.js";
export { OxmlWrapper, createElement, countAttrValues, xmlFragment, pfxmap } from "./oxml/base.js";
export { PackURI, PACKAGE_URI, CONTENT_TYPES_URI } from "./opc/packuri.js";
export {
  CONTENT_TYPE,
  RELATIONSHIP_TYPE,
  OPC_NAMESPACE,
  RELATIONSHIP_TARGET_MODE,
} from "./opc/constants.js";
export { defaultContentTypes, imageContentTypes } from "./opc/spec.js";
export {
  CT_Default,
  CT_Override,
  CT_Relationship,
  CT_Relationships,
  CT_Types,
  serializePartXml,
  XML_DECL,
} from "./opc/oxml.js";
export { PackageReader, PackageWriter } from "./opc/serialized.js";
export type { PackageItem } from "./opc/serialized.js";
export { ContentTypeMap } from "./opc/content-types.js";
export { Relationship, Relationships } from "./opc/rels.js";
export { OpcPackage, Part, XmlPart, PartFactory } from "./opc/package.js";
export type { PartConstructor } from "./opc/package.js";

// -- units & coordinates -------------------------------------------------
export {
  Emu,
  Inches,
  Pt,
  Cm,
  Mm,
  Pixels,
  Length,
  MAX_COLUMN,
  MAX_ROW,
  columnIndexFromString,
  getColumnLetter,
  coordinateFromString,
  cellRefToRowCol,
  rowColToCellRef,
  absoluteCoordinate,
  rangeBoundaries,
  boundariesToRange,
} from "./util.js";
export type { RangeBoundaries } from "./util.js";

// -- M1: workbook / worksheet / cell read + shared strings ---------------
export { BUILTIN_FORMATS, FORMAT_GENERAL, builtinFormatCode, isDateFormat } from "./numberFormats.js";
export { fromExcel, toExcel, WINDOWS_EPOCH_MS, MAC_EPOCH_MS } from "./datetimes.js";
export { CT_Workbook, CT_Sheet, successorsAfter } from "./oxml/workbook.js";
export { CT_Worksheet, CT_SheetData, CT_Row, CT_Cell } from "./oxml/worksheet.js";
export { CT_Sst, CT_Si } from "./oxml/sharedstrings.js";
export { CT_Stylesheet, CT_Xf } from "./oxml/styles.js";
export { WorkbookPart } from "./parts/workbook.js";
export { WorksheetPart } from "./parts/worksheet.js";
export { SharedStringsPart } from "./parts/sharedstrings.js";
export { StylesPart } from "./parts/styles.js";
export { Workbook } from "./workbook.js";
export type { WorkbookOpenOptions } from "./workbook.js";
export { Worksheet, ColumnDimension, RowDimension } from "./worksheet.js";
export type {
  CellRange,
  ConditionalRule,
  ConditionalFormatting,
  DataValidationSpec,
  HyperlinkSpec,
  CommentSpec,
} from "./worksheet.js";
export { Cell } from "./cell.js";
export type { CellValue, CellType } from "./cell.js";

// -- part-type registry (mirrors openpyxl's reader dispatch) -------------
import { PartFactory as _PartFactory } from "./opc/package.js";
import { CONTENT_TYPE as _CT } from "./opc/constants.js";
import { WorkbookPart as _WorkbookPart } from "./parts/workbook.js";
import { WorksheetPart as _WorksheetPart } from "./parts/worksheet.js";
import { SharedStringsPart as _SharedStringsPart } from "./parts/sharedstrings.js";
import { StylesPart as _StylesPart } from "./parts/styles.js";

_PartFactory.partTypeFor.set(_CT.SML_SHEET_MAIN, _WorkbookPart);
_PartFactory.partTypeFor.set(_CT.SML_TEMPLATE_MAIN, _WorkbookPart);
_PartFactory.partTypeFor.set(_CT.SML_SHEET_MACRO_ENABLED_MAIN, _WorkbookPart);
_PartFactory.partTypeFor.set(_CT.SML_TEMPLATE_MACRO_ENABLED_MAIN, _WorkbookPart);
_PartFactory.partTypeFor.set(_CT.SML_WORKSHEET, _WorksheetPart);
_PartFactory.partTypeFor.set(_CT.SML_SHARED_STRINGS, _SharedStringsPart);
_PartFactory.partTypeFor.set(_CT.SML_STYLES, _StylesPart);

// M6 part types
import { DrawingPart as _DrawingPart } from "./parts/drawing.js";
import { ChartPart as _ChartPart } from "./parts/chart.js";
import { ImagePart as _ImagePart } from "./parts/image.js";
import { CorePropertiesPart as _CorePropertiesPart } from "./parts/coreprops.js";
import { imageContentTypes as _imageContentTypes } from "./opc/spec.js";

_PartFactory.partTypeFor.set(_CT.SML_DRAWING, _DrawingPart);
_PartFactory.partTypeFor.set(_CT.DML_CHART, _ChartPart);
_PartFactory.partTypeFor.set(_CT.OPC_CORE_PROPERTIES, _CorePropertiesPart);
for (const ct of new Set(Object.values(_imageContentTypes))) {
  _PartFactory.partTypeFor.set(ct, _ImagePart);
}
_PartFactory.partTypeFor.set("image/jpg", _ImagePart);

// -- M3: styles (fonts / fills / borders / alignment / protection / numFmt) --
export {
  Color,
  Font,
  Side,
  Border,
  PatternFill,
  GradientFill,
  Alignment,
  Protection,
  fillFromElement,
} from "./styles/values.js";
export type {
  ColorOptions,
  FontOptions,
  UnderlineStyle,
  PatternType,
  PatternFillOptions,
  GradientFillOptions,
  GradientStop,
  BorderStyle,
  SideOptions,
  BorderOptions,
  AlignmentOptions,
  ProtectionOptions,
  Fill,
} from "./styles/values.js";
export { builtinFormatId } from "./numberFormats.js";
export { formatValue } from "./numberFormat.js";
export type { XfOverrides } from "./oxml/styles.js";

// -- M6: charts / images / core properties -------------------------------
export { probeImage } from "./image/probe.js";
export type { ImageProbe } from "./image/probe.js";
export { ImagePart, getOrAddImagePart } from "./parts/image.js";
export { DrawingPart } from "./parts/drawing.js";
export { ChartPart } from "./parts/chart.js";
export { parseChart } from "./oxml/chart.js";
export type { ChartInfo, SeriesInfo } from "./oxml/chart.js";
export { CoreProperties, CorePropertiesPart } from "./parts/coreprops.js";
