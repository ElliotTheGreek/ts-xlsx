import { XmlParseError } from "../exc.js";

const NAMED_ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  apos: "'",
  quot: '"',
};

/**
 * Decode XML character data: the five predefined entities plus decimal and
 * hexadecimal character references. Any other entity reference throws —
 * OOXML producers never emit them and a DTD would be required to define them.
 */
export function decodeXmlText(raw: string, baseOffset = 0): string {
  if (!raw.includes("&")) return raw;
  let out = "";
  let i = 0;
  for (;;) {
    const amp = raw.indexOf("&", i);
    if (amp === -1) {
      out += raw.slice(i);
      return out;
    }
    out += raw.slice(i, amp);
    const semi = raw.indexOf(";", amp + 1);
    if (semi === -1) {
      throw new XmlParseError("unterminated entity reference", baseOffset + amp);
    }
    const body = raw.slice(amp + 1, semi);
    if (body.startsWith("#x") || body.startsWith("#X")) {
      const digits = body.slice(2);
      if (!/^[0-9a-fA-F]+$/.test(digits)) {
        throw new XmlParseError(`invalid character reference &${body};`, baseOffset + amp);
      }
      out += codePointToString(parseInt(digits, 16), baseOffset + amp);
    } else if (body.startsWith("#")) {
      const digits = body.slice(1);
      if (!/^[0-9]+$/.test(digits)) {
        throw new XmlParseError(`invalid character reference &${body};`, baseOffset + amp);
      }
      out += codePointToString(parseInt(digits, 10), baseOffset + amp);
    } else {
      const ch = NAMED_ENTITIES[body];
      if (ch === undefined) {
        throw new XmlParseError(`unknown entity reference &${body};`, baseOffset + amp);
      }
      out += ch;
    }
    i = semi + 1;
  }
}

function codePointToString(cp: number, offset: number): string {
  if (!Number.isFinite(cp) || cp < 0 || cp > 0x10ffff || (cp >= 0xd800 && cp <= 0xdfff)) {
    throw new XmlParseError(`character reference out of range`, offset);
  }
  return String.fromCodePoint(cp);
}

/** Escape text content written through the API (raw source slices are never re-escaped). */
export function escapeXmlText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/**
 * Escape an attribute value written through the API. Tab, LF and CR are
 * emitted as character references so attribute-value normalization cannot
 * alter them on the next parse.
 */
export function escapeXmlAttr(s: string, quote: '"' | "'"): string {
  let out = s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/\t/g, "&#9;")
    .replace(/\n/g, "&#10;")
    .replace(/\r/g, "&#13;");
  if (quote === '"') out = out.replace(/"/g, "&quot;");
  else out = out.replace(/'/g, "&apos;");
  return out;
}
