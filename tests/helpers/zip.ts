import { readFileSync, readdirSync, existsSync } from "node:fs";
import { join } from "node:path";
import JSZip from "jszip";

const ROOT = join(import.meta.dirname, "..", "..");

/**
 * Every .xlsx available to tests: the authored oracle corpus in fixtures/ plus
 * real Excel-authored round-trip inputs in tests/assets/ (openpyxl-generated
 * styled/formula/chart/image workbooks; drop genuine Excel files here to
 * strengthen the corpus).
 */
export function allXlsxInputs(): { name: string; path: string }[] {
  const out: { name: string; path: string }[] = [];
  for (const dir of [join(ROOT, "fixtures"), join(ROOT, "tests", "assets")]) {
    if (!existsSync(dir)) continue;
    for (const f of readdirSync(dir)) {
      if (f.toLowerCase().endsWith(".xlsx")) out.push({ name: f, path: join(dir, f) });
    }
  }
  return out;
}

export function fixturePath(name: string): string {
  const p = join(ROOT, "fixtures", name);
  if (existsSync(p)) return p;
  return join(ROOT, "tests", "assets", name);
}

/** Decompressed entry map of a zip, in central-directory order. */
export async function zipEntries(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(bytes);
  const out = new Map<string, Uint8Array>();
  for (const name of Object.keys(zip.files)) {
    const f = zip.files[name]!;
    if (f.dir) continue;
    out.set(name, await f.async("uint8array"));
  }
  return out;
}

export async function zipEntriesOf(path: string): Promise<Map<string, Uint8Array>> {
  return zipEntries(new Uint8Array(readFileSync(path)));
}

export function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** First differing byte offset, with a short context window — for assertions. */
export function firstDiff(a: Uint8Array, b: Uint8Array): string {
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) {
    if (a[i] !== b[i]) {
      const dec = new TextDecoder();
      const ctx = (u: Uint8Array): string => dec.decode(u.subarray(Math.max(0, i - 40), i + 40));
      return `first diff at byte ${i}:\n  a: …${ctx(a)}…\n  b: …${ctx(b)}…`;
    }
  }
  return `lengths differ: ${a.length} vs ${b.length} (common prefix identical)`;
}
