/**
 * Number-format helpers — port of openpyxl/styles/numbers.py.
 *
 * Built-in format codes (ids < 164) plus the date-format heuristic openpyxl
 * uses to decide whether a numeric cell should be read back as a date/time.
 */

/** Standard format codes by id (ECMA-376 §18.8.30 + Excel built-ins). */
export const BUILTIN_FORMATS: Readonly<Record<number, string>> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: '"$"#,##0_);("$"#,##0)',
  6: '"$"#,##0_);[Red]("$"#,##0)',
  7: '"$"#,##0.00_);("$"#,##0.00)',
  8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "mm-dd-yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0_);(#,##0)",
  38: "#,##0_);[Red](#,##0)",
  39: "#,##0.00_);(#,##0.00)",
  40: "#,##0.00_);[Red](#,##0.00)",
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_)_("$"* \\(#,##0.00\\)_("$"* "-"??_)_(@_)',
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};

/** Custom (non-built-in) number formats start at this id. */
export const BUILTIN_FORMATS_MAX_SIZE = 164;

export const FORMAT_GENERAL = BUILTIN_FORMATS[0]!;

const BUILTIN_FORMATS_REVERSE: ReadonlyMap<string, number> = new Map(
  Object.entries(BUILTIN_FORMATS).map(([id, code]) => [code, Number(id)]),
);

/** Built-in format code for `id`, or undefined if `id` is not a built-in. */
export function builtinFormatCode(id: number): string | undefined {
  return BUILTIN_FORMATS[id];
}

/** Built-in numFmtId for a format code, or undefined when it is not a built-in. */
export function builtinFormatId(code: string): number | undefined {
  return BUILTIN_FORMATS_REVERSE.get(code);
}

// Strip quoted literals and [locale]/[color] groups (but not [h]/[m]/[s]
// elapsed-time markers) before scanning for date/time tokens — openpyxl STRIP_RE.
const LITERAL_GROUP = '".*?"';
const LOCALE_GROUP = "\\[(?!hh?\\]|mm?\\]|ss?\\])[^\\]]*\\]";
const STRIP_RE = new RegExp(`${LITERAL_GROUP}|${LOCALE_GROUP}`, "g");
const DATE_TOKEN_RE = /(?<![_\\])[dmhysDMHYS]/;

/**
 * True when a format code renders a date/time (openpyxl is_date_format): only
 * the first `;`-section is considered; quoted/locale groups are ignored; then a
 * bare d/m/h/y/s token (not escaped by `_` or `\`) marks it as a date format.
 */
export function isDateFormat(fmt: string | null | undefined): boolean {
  if (fmt == null) return false;
  let f = fmt.split(";")[0]!;
  f = f.replace(STRIP_RE, "");
  return DATE_TOKEN_RE.test(f);
}
