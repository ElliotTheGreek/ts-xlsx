import { XmlParseError } from "../exc.js";
import { XmlAttr, XmlCdata, XmlComment, XmlDocument, XmlElement, XmlPI, XmlText } from "./dom.js";

const NAME_START = /[A-Za-z_:À-￿]/;
const NAME_CHAR = /[A-Za-z0-9._:\-·À-￿]/;
const WS = /[ \t\r\n]/;

/**
 * Parse a well-formed UTF-8 XML document (no DOCTYPE, no external entities)
 * into the lossless DOM. Accepts bytes or an already-decoded string.
 */
export function parseXml(source: Uint8Array | string): XmlDocument {
  let text: string;
  let hadBom = false;
  if (typeof source === "string") {
    if (source.charCodeAt(0) === 0xfeff) {
      hadBom = true;
      text = source.slice(1);
    } else {
      text = source;
    }
  } else {
    if (source.length >= 2 && ((source[0] === 0xff && source[1] === 0xfe) || (source[0] === 0xfe && source[1] === 0xff))) {
      throw new XmlParseError("UTF-16 encoded XML is not supported (OOXML parts are UTF-8)", 0);
    }
    if (source.length >= 3 && source[0] === 0xef && source[1] === 0xbb && source[2] === 0xbf) {
      hadBom = true;
      source = source.subarray(3);
    }
    text = new TextDecoder("utf-8", { fatal: true }).decode(source);
  }
  return new Parser(text, hadBom).parseDocument();
}

class Parser {
  private pos = 0;

  constructor(
    private readonly s: string,
    private readonly hadBom: boolean,
  ) {}

  parseDocument(): XmlDocument {
    const declRaw = this.parseDeclaration();
    const prolog: (XmlComment | XmlPI | XmlText)[] = [];
    let root: XmlElement | null = null;

    while (this.pos < this.s.length && root === null) {
      const misc = this.parseMisc();
      if (misc) {
        prolog.push(misc);
        continue;
      }
      root = this.parseElement();
    }
    if (root === null) throw this.err("no root element");

    const epilog: (XmlComment | XmlPI | XmlText)[] = [];
    while (this.pos < this.s.length) {
      const misc = this.parseMisc();
      if (misc) epilog.push(misc);
      else throw this.err("content after root element");
    }

    const doc = new XmlDocument(root);
    doc.declRaw = declRaw;
    doc.hadBom = this.hadBom;
    doc.prolog = prolog;
    doc.epilog = epilog;
    return doc;
  }

  /** XML declaration, captured verbatim. */
  private parseDeclaration(): string | null {
    if (!this.s.startsWith("<?xml", this.pos)) return null;
    const after = this.s[this.pos + 5];
    if (after === undefined || !WS.test(after)) return null; // a PI like <?xmlfoo?>
    const end = this.s.indexOf("?>", this.pos);
    if (end === -1) throw this.err("unterminated XML declaration");
    const raw = this.s.slice(this.pos, end + 2);
    this.pos = end + 2;
    return raw;
  }

  /** Whitespace/comment/PI outside the root; null when the next thing is markup we don't own. */
  private parseMisc(): XmlComment | XmlPI | XmlText | null {
    const c = this.s[this.pos];
    if (c === undefined) return null;
    if (WS.test(c)) {
      const start = this.pos;
      while (this.pos < this.s.length && WS.test(this.s[this.pos]!)) this.pos++;
      return new XmlText(this.s.slice(start, this.pos));
    }
    if (this.s.startsWith("<!--", this.pos)) return this.parseComment();
    if (this.s.startsWith("<?", this.pos)) return this.parsePI();
    if (this.s.startsWith("<!DOCTYPE", this.pos)) {
      throw this.err("DOCTYPE is not supported (OOXML parts never declare one)");
    }
    if (c !== "<") throw this.err("text content outside root element");
    return null;
  }

  private parseComment(): XmlComment {
    const start = this.pos + 4;
    const end = this.s.indexOf("-->", start);
    if (end === -1) throw this.err("unterminated comment");
    this.pos = end + 3;
    return new XmlComment(this.s.slice(start, end));
  }

  private parsePI(): XmlPI {
    const start = this.pos;
    const end = this.s.indexOf("?>", start + 2);
    if (end === -1) throw this.err("unterminated processing instruction");
    const raw = this.s.slice(start, end + 2);
    this.pos = end + 2;
    const body = raw.slice(2, -2);
    const m = /^[^ \t\r\n?]+/.exec(body);
    return new XmlPI(raw, m ? m[0] : "");
  }

  private parseElement(): XmlElement {
    if (this.s[this.pos] !== "<") throw this.err("expected element");
    this.pos++;
    const name = this.parseName();

    const el = new XmlElement(name);
    // attributes
    for (;;) {
      const gapStart = this.pos;
      while (this.pos < this.s.length && WS.test(this.s[this.pos]!)) this.pos++;
      const gap = this.s.slice(gapStart, this.pos);
      const c = this.s[this.pos];
      if (c === undefined) throw this.err("unterminated start tag");
      if (c === ">") {
        el.endTagGap = gap;
        el.selfClosing = false;
        this.pos++;
        break;
      }
      if (c === "/") {
        if (this.s[this.pos + 1] !== ">") throw this.err("expected '/>'");
        el.endTagGap = gap;
        el.selfClosing = true;
        this.pos += 2;
        return el; // empty element — no content
      }
      if (gap === "") throw this.err("expected whitespace before attribute");
      el.loadAttr(this.parseAttr(gap));
    }

    // content
    for (;;) {
      if (this.pos >= this.s.length) throw this.err(`unterminated element <${name}>`);
      if (this.s.startsWith("</", this.pos)) {
        this.pos += 2;
        const closeName = this.parseName();
        while (this.pos < this.s.length && WS.test(this.s[this.pos]!)) this.pos++;
        if (this.s[this.pos] !== ">") throw this.err("malformed end tag");
        this.pos++;
        if (closeName !== name) {
          throw this.err(`end tag </${closeName}> does not match <${name}>`);
        }
        return el;
      }
      if (this.s.startsWith("<![CDATA[", this.pos)) {
        const start = this.pos + 9;
        const end = this.s.indexOf("]]>", start);
        if (end === -1) throw this.err("unterminated CDATA section");
        el.loadChild(new XmlCdata(this.s.slice(start, end)));
        this.pos = end + 3;
        continue;
      }
      if (this.s.startsWith("<!--", this.pos)) {
        el.loadChild(this.parseComment());
        continue;
      }
      if (this.s.startsWith("<?", this.pos)) {
        el.loadChild(this.parsePI());
        continue;
      }
      if (this.s[this.pos] === "<") {
        el.loadChild(this.parseElement());
        continue;
      }
      // character data — stored raw, decoded lazily
      const next = this.s.indexOf("<", this.pos);
      if (next === -1) throw this.err(`unterminated element <${name}>`);
      el.loadChild(new XmlText(this.s.slice(this.pos, next)));
      this.pos = next;
    }
  }

  private parseAttr(leadingGap: string): XmlAttr {
    const name = this.parseName();
    while (this.pos < this.s.length && WS.test(this.s[this.pos]!)) this.pos++;
    if (this.s[this.pos] !== "=") throw this.err(`expected '=' after attribute ${name}`);
    this.pos++;
    while (this.pos < this.s.length && WS.test(this.s[this.pos]!)) this.pos++;
    const quote = this.s[this.pos];
    if (quote !== '"' && quote !== "'") throw this.err(`expected quoted value for attribute ${name}`);
    this.pos++;
    const end = this.s.indexOf(quote, this.pos);
    if (end === -1) throw this.err(`unterminated value for attribute ${name}`);
    const rawValue = this.s.slice(this.pos, end);
    if (rawValue.includes("<")) throw this.err(`'<' in value of attribute ${name}`);
    this.pos = end + 1;
    return { name, rawValue, quote, leadingGap };
  }

  private parseName(): string {
    const start = this.pos;
    const first = this.s[this.pos];
    if (first === undefined || !NAME_START.test(first)) throw this.err("expected name");
    this.pos++;
    while (this.pos < this.s.length && NAME_CHAR.test(this.s[this.pos]!)) this.pos++;
    return this.s.slice(start, this.pos);
  }

  private err(message: string): XmlParseError {
    return new XmlParseError(message, this.pos);
  }
}
