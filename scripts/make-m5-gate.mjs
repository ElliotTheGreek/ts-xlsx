#!/usr/bin/env node
/** M5 gate — conditional formatting, data validation, hyperlink, defined name
 * for manual Excel review + openpyxl cross-validation. */
import { mkdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook, PatternFill } from "../dist/index.js";

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "..");
mkdirSync(join(root, "out"), { recursive: true });

const wb = await Workbook.open(join(root, "tests", "assets", "basic.xlsx"));
const ws = wb.get("Data");
ws.addCellIsRule("B2:B3", {
  operator: "greaterThan",
  formula: "20",
  fill: new PatternFill({ patternType: "solid", fgColor: "FFC7CE" }),
});
ws.addDataValidation({ sqref: "D1:D5", type: "list", formula1: '"x,y,z"', allowBlank: true });
ws.cell("A1").hyperlink = "https://flowdot.ai";
wb.setDefinedName("Scores", "Data!$B$2:$B$3");

const dest = join(root, "out", "m5-rules.xlsx");
await wb.save(dest);
console.log("wrote", dest);
