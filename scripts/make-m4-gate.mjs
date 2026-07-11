#!/usr/bin/env node
/** M4 gate — structural edits (merges, dims, insert/delete, freeze, sheet ops)
 * for manual Excel review + openpyxl cross-validation. */
import { mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook } from "../dist/index.js";

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "..");
mkdirSync(join(root, "out"), { recursive: true });

const wb = await Workbook.open(join(root, "tests", "assets", "basic.xlsx"));
const ws = wb.get("Data");
ws.mergeCells("A7:B8");
ws.column("B").width = 26;
ws.row(1).height = 24;
ws.freezePanes = "B2";
ws.autoFilter = "A1:C3";
ws.tabColor = "2196F3";
ws.insertRows(3, 1);

const fresh = wb.createSheet("Summary");
fresh.cell("A1").value = "Summary sheet";
fresh.cell("A2").value = 42;
wb.copyWorksheet(wb.get("Data"), "DataCopy");
wb.removeSheet("Hidden");
wb.moveSheet("Summary", 0);

const dest = join(root, "out", "m4-structure.xlsx");
await wb.save(dest);
console.log("wrote", dest);
