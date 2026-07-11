/**
 * Number-format display renderer — turn a cached cell value + its format code
 * into the string Excel would show. This is the "render a cached value for
 * display" helper the engine/AI needs; it is deliberately a *pragmatic* subset
 * of Excel's format mini-language (General, fixed/grouped decimals, percent,
 * currency literals, and date/time tokens). Exotic codes fall back to a
 * reasonable string rather than pretending full fidelity.
 */
import { fromExcel } from "./datetimes.js";
import { isDateFormat, FORMAT_GENERAL } from "./numberFormats.js";

const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

const pad = (n: number, w = 2): string => String(n).padStart(w, "0");

/** Render `value` per format `code`. `is1904` selects the date system for
 * serial-date values. */
export function formatValue(
  value: number | string | boolean | Date | null,
  code: string,
  is1904 = false,
): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (code === "" || code === FORMAT_GENERAL) return generalFormat(value);
  if (code === "@") return String(value);

  if (isDateFormat(code)) {
    const d = value instanceof Date ? value : fromExcel(Number(value), is1904);
    return renderDate(d, code.split(";")[0]!);
  }

  if (typeof value === "string") return value;
  return renderNumber(Number(value), code);
}

function generalFormat(value: number | string | Date): string {
  if (value instanceof Date) return value.toISOString();
  return String(value);
}

// -- numeric ------------------------------------------------------------------

function renderNumber(n: number, code: string): string {
  const sections = code.split(";");
  let section: string;
  if (n < 0 && sections.length > 1) section = sections[1]!;
  else if (n === 0 && sections.length > 2) section = sections[2]!;
  else section = sections[0]!;

  const percent = section.includes("%");
  let v = Math.abs(n);
  if (percent) v *= 100;

  const m = section.match(/[#0][#0,]*(\.[#0]+)?/);
  const placeholder = m ? m[0] : "0";
  const decimals = placeholder.includes(".")
    ? placeholder.split(".")[1]!.replace(/[^0#]/g, "").length
    : 0;
  const grouping = /[#0],[#0]/.test(placeholder) || placeholder.includes(",");

  let digits = v.toFixed(decimals);
  if (grouping) digits = groupThousands(digits);

  const idx = m ? section.indexOf(m[0]) : 0;
  const clean = (t: string): string =>
    t
      .replace(/\[[^\]]*\]/g, "")
      .replace(/"([^"]*)"/g, "$1")
      .replace(/\\(.)/g, "$1");
  const prefix = clean(section.slice(0, idx));
  const suffix = clean(section.slice(idx + placeholder.length));

  // A negative section (parentheses) carries its own sign; otherwise prepend "-".
  // The "%" literal, when present, is already part of `suffix`.
  const sign = n < 0 && sections.length < 2 ? "-" : "";
  return `${sign}${prefix}${digits}${suffix}`.trim();
}

function groupThousands(s: string): string {
  const [intPart, frac] = s.split(".");
  const grouped = intPart!.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return frac === undefined ? grouped : `${grouped}.${frac}`;
}

// -- date / time --------------------------------------------------------------

function renderDate(d: Date, section: string): string {
  const has12h = /AM\/PM/i.test(section);
  let out = "";
  let i = 0;
  // Track whether the previous time token was hours, so an "m" run means minutes.
  let afterHour = false;

  while (i < section.length) {
    const ch = section[i]!;

    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      out += section.slice(i + 1, end === -1 ? section.length : end);
      i = end === -1 ? section.length : end + 1;
      continue;
    }
    if (ch === "\\") {
      out += section[i + 1] ?? "";
      i += 2;
      continue;
    }
    if (ch === "[") {
      // elapsed-time markers like [h] — treat [hh] as hours here (pragmatic)
      const end = section.indexOf("]", i);
      const inner = section.slice(i + 1, end);
      if (/^h+$/i.test(inner)) out += pad(d.getUTCHours(), inner.length);
      i = end === -1 ? section.length : end + 1;
      continue;
    }

    const lower = ch.toLowerCase();
    if ("ymdhs".includes(lower)) {
      let j = i;
      while (j < section.length && section[j]!.toLowerCase() === lower) j++;
      const run = section.slice(i, j);
      const len = run.length;

      if (lower === "y") {
        out += len <= 2 ? pad(d.getUTCFullYear() % 100) : String(d.getUTCFullYear());
      } else if (lower === "d") {
        if (len === 1) out += String(d.getUTCDate());
        else if (len === 2) out += pad(d.getUTCDate());
        else if (len === 3) out += DAYS[d.getUTCDay()]!.slice(0, 3);
        else out += DAYS[d.getUTCDay()];
      } else if (lower === "h") {
        const h24 = d.getUTCHours();
        const h = has12h ? (h24 % 12 === 0 ? 12 : h24 % 12) : h24;
        out += len === 1 ? String(h) : pad(h);
        afterHour = true;
      } else if (lower === "s") {
        out += len === 1 ? String(d.getUTCSeconds()) : pad(d.getUTCSeconds());
      } else if (lower === "m") {
        // month vs minute: minutes when it follows an hour token or precedes seconds
        const rest = section.slice(j).replace(/[^a-zA-Z]/g, "");
        const isMinute = afterHour || /^s/i.test(rest);
        if (isMinute) {
          out += len === 1 ? String(d.getUTCMinutes()) : pad(d.getUTCMinutes());
        } else if (len === 1) out += String(d.getUTCMonth() + 1);
        else if (len === 2) out += pad(d.getUTCMonth() + 1);
        else if (len === 3) out += MONTHS[d.getUTCMonth()]!.slice(0, 3);
        else out += MONTHS[d.getUTCMonth()];
      }
      if (lower !== "h") afterHour = false;
      i = j;
      continue;
    }

    if (/^am\/pm$/i.test(section.slice(i, i + 5))) {
      out += d.getUTCHours() < 12 ? "AM" : "PM";
      i += 5;
      continue;
    }

    out += ch;
    i++;
  }
  return out;
}
