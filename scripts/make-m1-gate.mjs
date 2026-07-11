#!/usr/bin/env node
/**
 * M0/M1 gate: open every available .xlsx with zero edits and save a copy to
 * out/. Manual checklist (Elliot): each output opens in Excel with no repair
 * prompt, looks identical, and Excel can re-save it. Proves the fidelity
 * backbone (untouched parts byte-identical, charts/images/orphans preserved).
 *
 * Run: npm run build && node scripts/make-m1-gate.mjs
 */
import { mkdirSync, readdirSync, existsSync } from "node:fs";
import { join, dirname, basename } from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook } from "../dist/index.js";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const outDir = join(root, "out");
mkdirSync(outDir, { recursive: true });

const inputs = [];
for (const dir of ["fixtures", join("tests", "assets")]) {
  const abs = join(root, dir);
  if (!existsSync(abs)) continue;
  for (const f of readdirSync(abs)) {
    if (f.toLowerCase().endsWith(".xlsx")) inputs.push(join(abs, f));
  }
}

for (const path of inputs) {
  const wb = await Workbook.open(path);
  // touch the read surface (must not mutate), then save a zero-edit copy
  for (const ws of wb.worksheets) void ws.dimensions;
  const dest = join(outDir, `m1-roundtrip-${basename(path)}`);
  await wb.save(dest);
  console.log("wrote", dest);
}
