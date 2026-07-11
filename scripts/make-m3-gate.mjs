#!/usr/bin/env node
/** M3 gate — apply fonts/fills/borders/alignment/number formats and save for
 * manual Excel review + openpyxl cross-validation. */
import { mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook, Font, PatternFill, Border, Side, Alignment } from "../dist/index.js";

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "..");
mkdirSync(join(root, "out"), { recursive: true });

const wb = await Workbook.open(join(root, "tests", "assets", "basic.xlsx"));
const ws = wb.get("Data");
ws.cell("A2").font = new Font({ bold: true, italic: true, color: "0000FF", size: 13 });
ws.cell("B3").fill = new PatternFill({ patternType: "solid", fgColor: "FFEB3B" });
ws.cell("C3").border = new Border({
  top: new Side({ style: "medium", color: "FF0000" }),
  bottom: new Side({ style: "double" }),
});
ws.cell("A3").alignment = new Alignment({ horizontal: "right", vertical: "center", wrapText: true });
ws.cell("B2").numberFormat = "0.00%";
ws.cell("A1").numberFormat = "#,##0.000";

const dest = join(root, "out", "m3-styles.xlsx");
await wb.save(dest);
console.log("wrote", dest);
