/**
 * Style value objects — port of openpyxl/styles/{colors,fonts,fills,borders,
 * alignment,protection}.py. Each class is an immutable-ish data holder that
 * knows its own SpreadsheetML shape: `fromElement` parses, `writeInto`
 * populates a fresh element, and `key()` yields a canonical identity used by
 * the stylesheet's find-or-add dedup so the tables never bloat.
 *
 * Reads never mutate: parsing an element constructs a new object and leaves the
 * XML untouched; a Cell style setter only ever writes when you assign.
 */
import { XmlElement } from "../xml/dom.js";
import { createElement } from "../oxml/base.js";
import { nsmap, NsTag } from "../oxml/ns.js";

const MAIN = nsmap.main;

function child(parent: XmlElement, tag: NsTag): XmlElement {
  const el = createElement(tag, parent);
  parent.appendChild(el);
  return el;
}

function boolAttr(el: XmlElement, name: string): boolean | undefined {
  const v = el.getAttr(name);
  if (v === null) return undefined;
  return v === "1" || v === "true";
}

// -- Color --------------------------------------------------------------------

export interface ColorOptions {
  rgb?: string;
  theme?: number;
  indexed?: number;
  tint?: number;
  auto?: boolean;
}

/** An ARGB / theme / indexed color reference (openpyxl Color). */
export class Color {
  readonly rgb?: string;
  readonly theme?: number;
  readonly indexed?: number;
  readonly tint?: number;
  readonly auto?: boolean;

  constructor(opts: ColorOptions | string = {}) {
    if (typeof opts === "string") {
      this.rgb = normalizeRgb(opts);
    } else {
      if (opts.rgb !== undefined) this.rgb = normalizeRgb(opts.rgb);
      this.theme = opts.theme;
      this.indexed = opts.indexed;
      this.tint = opts.tint;
      this.auto = opts.auto;
    }
  }

  static fromElement(el: XmlElement | null): Color | undefined {
    if (el === null) return undefined;
    const rgb = el.getAttr("rgb");
    const theme = el.getAttr("theme");
    const indexed = el.getAttr("indexed");
    const tint = el.getAttr("tint");
    const auto = el.getAttr("auto");
    return new Color({
      rgb: rgb ?? undefined,
      theme: theme === null ? undefined : Number(theme),
      indexed: indexed === null ? undefined : Number(indexed),
      tint: tint === null ? undefined : Number(tint),
      auto: auto === null ? undefined : auto === "1" || auto === "true",
    });
  }

  writeInto(el: XmlElement): void {
    if (this.rgb !== undefined) el.setAttr("rgb", this.rgb);
    if (this.theme !== undefined) el.setAttr("theme", String(this.theme));
    if (this.indexed !== undefined) el.setAttr("indexed", String(this.indexed));
    if (this.tint !== undefined) el.setAttr("tint", String(this.tint));
    if (this.auto !== undefined) el.setAttr("auto", this.auto ? "1" : "0");
  }

  key(): string {
    return `rgb=${this.rgb ?? ""};th=${this.theme ?? ""};ix=${this.indexed ?? ""};ti=${this.tint ?? ""};au=${this.auto ?? ""}`;
  }
}

function normalizeRgb(s: string): string {
  const up = s.toUpperCase();
  if (/^[0-9A-F]{8}$/.test(up)) return up;
  if (/^[0-9A-F]{6}$/.test(up)) return `00${up}`; // openpyxl prepends alpha 00
  throw new Error(`invalid rgb color: "${s}" (expected 6 or 8 hex digits)`);
}

// -- Font ---------------------------------------------------------------------

export type UnderlineStyle = "single" | "double" | "singleAccounting" | "doubleAccounting";

export interface FontOptions {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: UnderlineStyle | boolean;
  strike?: boolean;
  color?: Color | string;
  vertAlign?: "superscript" | "subscript" | "baseline";
  family?: number;
  scheme?: "major" | "minor";
}

export class Font {
  readonly name?: string;
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: UnderlineStyle;
  readonly strike?: boolean;
  readonly color?: Color;
  readonly vertAlign?: "superscript" | "subscript" | "baseline";
  readonly family?: number;
  readonly scheme?: "major" | "minor";

  constructor(opts: FontOptions = {}) {
    this.name = opts.name;
    this.size = opts.size;
    this.bold = opts.bold;
    this.italic = opts.italic;
    this.underline =
      opts.underline === true ? "single" : opts.underline === false ? undefined : opts.underline;
    this.strike = opts.strike;
    this.color = opts.color === undefined ? undefined : toColor(opts.color);
    this.vertAlign = opts.vertAlign;
    this.family = opts.family;
    this.scheme = opts.scheme;
  }

  static fromElement(el: XmlElement): Font {
    const valOf = (ln: string): string | undefined => el.find(MAIN, ln)?.getAttr("val") ?? undefined;
    const flag = (ln: string): boolean | undefined => {
      const f = el.find(MAIN, ln);
      if (f === null) return undefined;
      const v = f.getAttr("val");
      return v === null ? true : v === "1" || v === "true";
    };
    const sz = valOf("sz");
    const family = valOf("family");
    const u = el.find(MAIN, "u");
    return new Font({
      name: valOf("name"),
      size: sz === undefined ? undefined : Number(sz),
      bold: flag("b"),
      italic: flag("i"),
      underline: u === null ? undefined : ((u.getAttr("val") ?? "single") as UnderlineStyle),
      strike: flag("strike"),
      color: Color.fromElement(el.find(MAIN, "color")),
      vertAlign: el.find(MAIN, "vertAlign")?.getAttr("val") as Font["vertAlign"],
      family: family === undefined ? undefined : Number(family),
      scheme: valOf("scheme") as Font["scheme"],
    });
  }

  /** Populate a fresh `<font>` element (child order per ECMA-376 §18.8.22). */
  writeInto(el: XmlElement): void {
    const setVal = (tag: NsTag, val: string): void => child(el, tag).setAttr("val", val);
    const setFlag = (tag: NsTag): void => child(el, tag).setAttr("val", "1");
    if (this.bold) setFlag("main:b");
    if (this.italic) setFlag("main:i");
    if (this.strike) setFlag("main:strike");
    if (this.underline !== undefined) {
      if (this.underline === "single") child(el, "main:u");
      else setVal("main:u", this.underline);
    }
    if (this.vertAlign !== undefined) setVal("main:vertAlign", this.vertAlign);
    if (this.size !== undefined) setVal("main:sz", numStr(this.size));
    if (this.color !== undefined) this.color.writeInto(child(el, "main:color"));
    if (this.name !== undefined) setVal("main:name", this.name);
    if (this.family !== undefined) setVal("main:family", String(this.family));
    if (this.scheme !== undefined) setVal("main:scheme", this.scheme);
  }

  key(): string {
    return [
      this.name ?? "",
      this.size ?? "",
      this.bold ? 1 : 0,
      this.italic ? 1 : 0,
      this.underline ?? "",
      this.strike ? 1 : 0,
      this.color?.key() ?? "",
      this.vertAlign ?? "",
      this.family ?? "",
      this.scheme ?? "",
    ].join("|");
  }
}

// -- Fills --------------------------------------------------------------------

export type PatternType =
  | "none"
  | "solid"
  | "gray125"
  | "darkGray"
  | "mediumGray"
  | "lightGray"
  | "gray0625"
  | "darkHorizontal"
  | "darkVertical"
  | "darkDown"
  | "darkUp"
  | "darkGrid"
  | "darkTrellis"
  | "lightHorizontal"
  | "lightVertical"
  | "lightDown"
  | "lightUp"
  | "lightGrid"
  | "lightTrellis";

export interface PatternFillOptions {
  patternType?: PatternType;
  fgColor?: Color | string;
  bgColor?: Color | string;
}

export class PatternFill {
  readonly patternType?: PatternType;
  readonly fgColor?: Color;
  readonly bgColor?: Color;

  constructor(opts: PatternFillOptions = {}) {
    this.patternType = opts.patternType;
    this.fgColor = opts.fgColor === undefined ? undefined : toColor(opts.fgColor);
    this.bgColor = opts.bgColor === undefined ? undefined : toColor(opts.bgColor);
  }

  static fromElement(el: XmlElement): PatternFill {
    const pt = el.getAttr("patternType");
    return new PatternFill({
      patternType: (pt ?? undefined) as PatternType | undefined,
      fgColor: Color.fromElement(el.find(MAIN, "fgColor")),
      bgColor: Color.fromElement(el.find(MAIN, "bgColor")),
    });
  }

  writeInto(fillEl: XmlElement): void {
    const pf = child(fillEl, "main:patternFill");
    if (this.patternType !== undefined) pf.setAttr("patternType", this.patternType);
    if (this.fgColor !== undefined) this.fgColor.writeInto(child(pf, "main:fgColor"));
    if (this.bgColor !== undefined) this.bgColor.writeInto(child(pf, "main:bgColor"));
  }

  key(): string {
    return `pat|${this.patternType ?? ""}|${this.fgColor?.key() ?? ""}|${this.bgColor?.key() ?? ""}`;
  }
}

export interface GradientStop {
  position: number;
  color: Color;
}

export interface GradientFillOptions {
  type?: "linear" | "path";
  degree?: number;
  stops?: GradientStop[];
}

export class GradientFill {
  readonly type: "linear" | "path";
  readonly degree: number;
  readonly stops: GradientStop[];

  constructor(opts: GradientFillOptions = {}) {
    this.type = opts.type ?? "linear";
    this.degree = opts.degree ?? 0;
    this.stops = opts.stops ?? [];
  }

  static fromElement(el: XmlElement): GradientFill {
    const stops: GradientStop[] = [];
    for (const s of el.findAll(MAIN, "stop")) {
      const pos = s.getAttr("position");
      stops.push({
        position: pos === null ? 0 : Number(pos),
        color: Color.fromElement(s.find(MAIN, "color")) ?? new Color(),
      });
    }
    return new GradientFill({
      type: (el.getAttr("type") ?? "linear") as "linear" | "path",
      degree: el.getAttr("degree") === null ? 0 : Number(el.getAttr("degree")),
      stops,
    });
  }

  writeInto(fillEl: XmlElement): void {
    const gf = child(fillEl, "main:gradientFill");
    if (this.type !== "linear") gf.setAttr("type", this.type);
    if (this.degree !== 0) gf.setAttr("degree", numStr(this.degree));
    for (const s of this.stops) {
      const stop = child(gf, "main:stop");
      stop.setAttr("position", numStr(s.position));
      s.color.writeInto(child(stop, "main:color"));
    }
  }

  key(): string {
    return `grad|${this.type}|${this.degree}|${this.stops.map((s) => `${s.position}:${s.color.key()}`).join(",")}`;
  }
}

export type Fill = PatternFill | GradientFill;

/** Parse a `<fill>` into a PatternFill or GradientFill. */
export function fillFromElement(fillEl: XmlElement): Fill {
  const grad = fillEl.find(MAIN, "gradientFill");
  if (grad !== null) return GradientFill.fromElement(grad);
  const pat = fillEl.find(MAIN, "patternFill");
  return pat === null ? new PatternFill() : PatternFill.fromElement(pat);
}

// -- Borders ------------------------------------------------------------------

export type BorderStyle =
  | "thin"
  | "medium"
  | "thick"
  | "dashed"
  | "dotted"
  | "double"
  | "hair"
  | "mediumDashed"
  | "dashDot"
  | "mediumDashDot"
  | "dashDotDot"
  | "mediumDashDotDot"
  | "slantDashDot";

export interface SideOptions {
  style?: BorderStyle;
  color?: Color | string;
}

export class Side {
  readonly style?: BorderStyle;
  readonly color?: Color;

  constructor(opts: SideOptions = {}) {
    this.style = opts.style;
    this.color = opts.color === undefined ? undefined : toColor(opts.color);
  }

  static fromElement(el: XmlElement | null): Side {
    if (el === null) return new Side();
    return new Side({
      style: (el.getAttr("style") ?? undefined) as BorderStyle | undefined,
      color: Color.fromElement(el.find(MAIN, "color")),
    });
  }

  writeInto(el: XmlElement): void {
    if (this.style !== undefined) el.setAttr("style", this.style);
    if (this.color !== undefined) this.color.writeInto(child(el, "main:color"));
  }

  key(): string {
    return `${this.style ?? ""}:${this.color?.key() ?? ""}`;
  }
}

export interface BorderOptions {
  left?: Side;
  right?: Side;
  top?: Side;
  bottom?: Side;
  diagonal?: Side;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
}

export class Border {
  readonly left: Side;
  readonly right: Side;
  readonly top: Side;
  readonly bottom: Side;
  readonly diagonal: Side;
  readonly diagonalUp?: boolean;
  readonly diagonalDown?: boolean;

  constructor(opts: BorderOptions = {}) {
    this.left = opts.left ?? new Side();
    this.right = opts.right ?? new Side();
    this.top = opts.top ?? new Side();
    this.bottom = opts.bottom ?? new Side();
    this.diagonal = opts.diagonal ?? new Side();
    this.diagonalUp = opts.diagonalUp;
    this.diagonalDown = opts.diagonalDown;
  }

  static fromElement(el: XmlElement): Border {
    return new Border({
      left: Side.fromElement(el.find(MAIN, "left") ?? el.find(MAIN, "start")),
      right: Side.fromElement(el.find(MAIN, "right") ?? el.find(MAIN, "end")),
      top: Side.fromElement(el.find(MAIN, "top")),
      bottom: Side.fromElement(el.find(MAIN, "bottom")),
      diagonal: Side.fromElement(el.find(MAIN, "diagonal")),
      diagonalUp: boolAttr(el, "diagonalUp"),
      diagonalDown: boolAttr(el, "diagonalDown"),
    });
  }

  /** Populate a fresh `<border>` (child order: left,right,top,bottom,diagonal). */
  writeInto(el: XmlElement): void {
    if (this.diagonalUp !== undefined) el.setAttr("diagonalUp", this.diagonalUp ? "1" : "0");
    if (this.diagonalDown !== undefined) el.setAttr("diagonalDown", this.diagonalDown ? "1" : "0");
    this.left.writeInto(child(el, "main:left"));
    this.right.writeInto(child(el, "main:right"));
    this.top.writeInto(child(el, "main:top"));
    this.bottom.writeInto(child(el, "main:bottom"));
    this.diagonal.writeInto(child(el, "main:diagonal"));
  }

  key(): string {
    return [
      this.left.key(),
      this.right.key(),
      this.top.key(),
      this.bottom.key(),
      this.diagonal.key(),
      this.diagonalUp ? 1 : 0,
      this.diagonalDown ? 1 : 0,
    ].join("|");
  }
}

// -- Alignment ----------------------------------------------------------------

export interface AlignmentOptions {
  horizontal?: string;
  vertical?: string;
  textRotation?: number;
  wrapText?: boolean;
  shrinkToFit?: boolean;
  indent?: number;
}

export class Alignment {
  readonly horizontal?: string;
  readonly vertical?: string;
  readonly textRotation?: number;
  readonly wrapText?: boolean;
  readonly shrinkToFit?: boolean;
  readonly indent?: number;

  constructor(opts: AlignmentOptions = {}) {
    this.horizontal = opts.horizontal;
    this.vertical = opts.vertical;
    this.textRotation = opts.textRotation;
    this.wrapText = opts.wrapText;
    this.shrinkToFit = opts.shrinkToFit;
    this.indent = opts.indent;
  }

  static fromElement(el: XmlElement | null): Alignment {
    if (el === null) return new Alignment();
    const num = (n: string): number | undefined => {
      const v = el.getAttr(n);
      return v === null ? undefined : Number(v);
    };
    return new Alignment({
      horizontal: el.getAttr("horizontal") ?? undefined,
      vertical: el.getAttr("vertical") ?? undefined,
      textRotation: num("textRotation"),
      wrapText: boolAttr(el, "wrapText"),
      shrinkToFit: boolAttr(el, "shrinkToFit"),
      indent: num("indent"),
    });
  }

  writeInto(el: XmlElement): void {
    if (this.horizontal !== undefined) el.setAttr("horizontal", this.horizontal);
    if (this.vertical !== undefined) el.setAttr("vertical", this.vertical);
    if (this.textRotation !== undefined) el.setAttr("textRotation", String(this.textRotation));
    if (this.wrapText !== undefined) el.setAttr("wrapText", this.wrapText ? "1" : "0");
    if (this.shrinkToFit !== undefined) el.setAttr("shrinkToFit", this.shrinkToFit ? "1" : "0");
    if (this.indent !== undefined) el.setAttr("indent", String(this.indent));
  }

  get isEmpty(): boolean {
    return (
      this.horizontal === undefined &&
      this.vertical === undefined &&
      this.textRotation === undefined &&
      this.wrapText === undefined &&
      this.shrinkToFit === undefined &&
      this.indent === undefined
    );
  }

  key(): string {
    return [
      this.horizontal ?? "",
      this.vertical ?? "",
      this.textRotation ?? "",
      this.wrapText ? 1 : 0,
      this.shrinkToFit ? 1 : 0,
      this.indent ?? "",
    ].join("|");
  }
}

// -- Protection ---------------------------------------------------------------

export interface ProtectionOptions {
  locked?: boolean;
  hidden?: boolean;
}

export class Protection {
  readonly locked?: boolean;
  readonly hidden?: boolean;

  constructor(opts: ProtectionOptions = {}) {
    this.locked = opts.locked;
    this.hidden = opts.hidden;
  }

  static fromElement(el: XmlElement | null): Protection {
    if (el === null) return new Protection();
    return new Protection({ locked: boolAttr(el, "locked"), hidden: boolAttr(el, "hidden") });
  }

  writeInto(el: XmlElement): void {
    if (this.locked !== undefined) el.setAttr("locked", this.locked ? "1" : "0");
    if (this.hidden !== undefined) el.setAttr("hidden", this.hidden ? "1" : "0");
  }

  get isEmpty(): boolean {
    return this.locked === undefined && this.hidden === undefined;
  }

  key(): string {
    return `${this.locked ?? ""}|${this.hidden ?? ""}`;
  }
}

// -- helpers ------------------------------------------------------------------

function toColor(c: Color | string): Color {
  return c instanceof Color ? c : new Color(c);
}

function numStr(n: number): string {
  return Number.isInteger(n) ? String(n) : String(n);
}
