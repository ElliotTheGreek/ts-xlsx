/**
 * OPC constants — SpreadsheetML content types and relationship types
 * (ECMA-376 / openpyxl xml/constants.py). Values verified against the ISO
 * spec and openpyxl source.
 */

export const CONTENT_TYPE = {
  // -- images (Default entries) --
  BMP: "image/bmp",
  GIF: "image/gif",
  JPEG: "image/jpeg",
  MS_PHOTO: "image/vnd.ms-photo",
  PNG: "image/png",
  TIFF: "image/tiff",
  X_EMF: "image/x-emf",
  X_WMF: "image/x-wmf",

  // -- package / office-wide --
  OPC_CORE_PROPERTIES: "application/vnd.openxmlformats-package.core-properties+xml",
  OPC_RELATIONSHIPS: "application/vnd.openxmlformats-package.relationships+xml",
  OFC_EXTENDED_PROPERTIES:
    "application/vnd.openxmlformats-officedocument.extended-properties+xml",
  OFC_CUSTOM_PROPERTIES:
    "application/vnd.openxmlformats-officedocument.custom-properties+xml",
  OFC_THEME: "application/vnd.openxmlformats-officedocument.theme+xml",
  OFC_VML_DRAWING: "application/vnd.openxmlformats-officedocument.vmlDrawing",
  XML: "application/xml",

  // -- SpreadsheetML workbook variants --
  SML_SHEET_MAIN: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  SML_TEMPLATE_MAIN:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
  SML_SHEET_MACRO_ENABLED_MAIN: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
  SML_TEMPLATE_MACRO_ENABLED_MAIN: "application/vnd.ms-excel.template.macroEnabled.main+xml",

  // -- SpreadsheetML parts --
  SML_WORKSHEET: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
  SML_CHARTSHEET: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
  SML_DIALOGSHEET: "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
  SML_SHARED_STRINGS:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
  SML_STYLES: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
  SML_CALC_CHAIN: "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
  SML_COMMENTS: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
  SML_TABLE: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
  SML_PIVOT_TABLE: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
  SML_PIVOT_CACHE_DEFINITION:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
  SML_PIVOT_CACHE_RECORDS:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
  SML_QUERY_TABLE: "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
  SML_EXTERNAL_LINK:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
  SML_SHEET_METADATA:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
  SML_VOLATILE_DEPENDENCIES:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml",
  SML_CONNECTIONS: "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
  SML_USER_NAMES: "application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml",
  SML_REVISION_HEADERS:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml",
  SML_REVISION_LOG:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml",
  SML_PRINTER_SETTINGS:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings",

  // -- DrawingML (charts + drawings embedded in sheets) --
  DML_CHART: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
  DML_CHARTSHAPES: "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml",
  DML_CHART_STYLE: "application/vnd.ms-office.chartstyle+xml",
  DML_CHART_COLOR_STYLE: "application/vnd.ms-office.chartcolorstyle+xml",
  OFC_CHART_EX: "application/vnd.ms-office.chartex+xml",
  SML_DRAWING: "application/vnd.openxmlformats-officedocument.drawing+xml",
} as const;

const REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const RELATIONSHIP_TYPE = {
  OFFICE_DOCUMENT: `${REL_BASE}/officeDocument`,
  WORKSHEET: `${REL_BASE}/worksheet`,
  CHARTSHEET: `${REL_BASE}/chartsheet`,
  DIALOGSHEET: `${REL_BASE}/dialogsheet`,
  SHARED_STRINGS: `${REL_BASE}/sharedStrings`,
  STYLES: `${REL_BASE}/styles`,
  THEME: `${REL_BASE}/theme`,
  CALC_CHAIN: `${REL_BASE}/calcChain`,
  COMMENTS: `${REL_BASE}/comments`,
  VML_DRAWING: `${REL_BASE}/vmlDrawing`,
  TABLE: `${REL_BASE}/table`,
  PIVOT_TABLE: `${REL_BASE}/pivotTable`,
  PIVOT_CACHE_DEFINITION: `${REL_BASE}/pivotCacheDefinition`,
  PIVOT_CACHE_RECORDS: `${REL_BASE}/pivotCacheRecords`,
  QUERY_TABLE: `${REL_BASE}/queryTable`,
  EXTERNAL_LINK: `${REL_BASE}/externalLink`,
  EXTERNAL_LINK_PATH: `${REL_BASE}/externalLinkPath`,
  SHEET_METADATA: `${REL_BASE}/sheetMetadata`,
  VOLATILE_DEPENDENCIES: `${REL_BASE}/volatileDependencies`,
  CONNECTIONS: `${REL_BASE}/connections`,
  DRAWING: `${REL_BASE}/drawing`,
  CHART: `${REL_BASE}/chart`,
  CHART_USER_SHAPES: `${REL_BASE}/chartUserShapes`,
  IMAGE: `${REL_BASE}/image`,
  HYPERLINK: `${REL_BASE}/hyperlink`,
  PRINTER_SETTINGS: `${REL_BASE}/printerSettings`,
  EXTENDED_PROPERTIES: `${REL_BASE}/extended-properties`,
  CUSTOM_PROPERTIES: `${REL_BASE}/custom-properties`,
  CUSTOM_XML: `${REL_BASE}/customXml`,
  CORE_PROPERTIES:
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
} as const;

/** Namespace of the OPC relationships and content-types XML vocabularies. */
export const OPC_NAMESPACE = {
  RELATIONSHIPS: "http://schemas.openxmlformats.org/package/2006/relationships",
  CONTENT_TYPES: "http://schemas.openxmlformats.org/package/2006/content-types",
} as const;

export const RELATIONSHIP_TARGET_MODE = {
  INTERNAL: "Internal",
  EXTERNAL: "External",
} as const;
