#!/usr/bin/env node
/**
 * M2 gate — edit an existing workbook and save it for manual Excel review +
 * openpyxl cross-validation. Proves ts-xlsx output opens cleanly and reads back
 * with the values/types Excel expects.
 */
import { mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook } from "../dist/index.js";

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "..");
const out = join(root, "out");
mkdirSync(out, { recursive: true });

const wb = await Workbook.open(join(root, "tests", "assets", "basic.xlsx"));
const ws = wb.get("Data");
ws.cell("A2").value = "Alice (edited)";
ws.cell("B2").value = 1234.56;
ws.cell("D1").value = true;
ws.cell("D2").value = "shared string";
ws.cell("D3").value = "shared string"; // dedupes with D2
ws.cell("D4").value = "=B2*2";
ws.cell("D5").value = new Date(Date.UTC(2030, 5, 15, 13, 45, 0));
ws.cell("D6").value = new Date(Date.UTC(2030, 0, 1));

const dest = join(out, "m2-edits.xlsx");
await wb.save(dest);
console.log("wrote", dest);
