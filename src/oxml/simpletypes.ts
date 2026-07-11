/**
 * Simple-type converters — port of pptx/oxml/simpletypes.py (the subset the
 * foundation needs; later milestones add theirs alongside).
 *
 * A SimpleType maps between the XML attribute string form and the TS value.
 */
import { InvalidXmlError } from "../exc.js";

export interface SimpleType<T> {
  fromXml(s: string): T;
  toXml(v: T): string;
}

export const XsdString: SimpleType<string> = {
  fromXml: (s) => s,
  toXml: (v) => v,
};

export const XsdId: SimpleType<string> = XsdString;
export const XsdAnyUri: SimpleType<string> = XsdString;

export const XsdBoolean: SimpleType<boolean> = {
  fromXml(s) {
    if (s === "1" || s === "true") return true;
    if (s === "0" || s === "false") return false;
    throw new InvalidXmlError(`expected xsd:boolean, got "${s}"`);
  },
  toXml: (v) => (v ? "1" : "0"),
};

function intType(min: number, max: number, name: string): SimpleType<number> {
  return {
    fromXml(s) {
      const v = Number(s);
      if (!Number.isInteger(v) || v < min || v > max) {
        throw new InvalidXmlError(`expected ${name}, got "${s}"`);
      }
      return v;
    },
    toXml(v) {
      if (!Number.isInteger(v) || v < min || v > max) {
        throw new InvalidXmlError(`value ${v} out of range for ${name}`);
      }
      return String(v);
    },
  };
}

export const XsdInt = intType(-2147483648, 2147483647, "xsd:int");
export const XsdUnsignedInt = intType(0, 4294967295, "xsd:unsignedInt");
// JS numbers are exact to 2^53-1; EMU values never approach that.
export const XsdLong = intType(Number.MIN_SAFE_INTEGER, Number.MAX_SAFE_INTEGER, "xsd:long");

/** OPC TargetMode attribute: "Internal" | "External". */
export const ST_TargetMode: SimpleType<"Internal" | "External"> = {
  fromXml(s) {
    if (s === "Internal" || s === "External") return s;
    throw new InvalidXmlError(`expected TargetMode, got "${s}"`);
  },
  toXml: (v) => v,
};
