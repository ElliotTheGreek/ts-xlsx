/** Base class for all ts-xlsx errors. */
export class XlsxError extends Error {
  constructor(message: string) {
    super(message);
    this.name = new.target.name;
  }
}

/** Raised when the source is not a valid OPC/zip package. */
export class PackageNotFoundError extends XlsxError {}

/** Raised when XML content violates an OOXML schema expectation. */
export class InvalidXmlError extends XlsxError {}

/** Raised when a part's XML is not well-formed. Carries the character offset. */
export class XmlParseError extends XlsxError {
  readonly offset: number;
  constructor(message: string, offset: number) {
    super(`${message} (at offset ${offset})`);
    this.offset = offset;
  }
}

/** Raised when a lookup (rId, reltype, partname, sheet name...) finds nothing. */
export class KeyLookupError extends XlsxError {}
