#!/usr/bin/env node
/** M6 gate — add an image, set core properties, and confirm charts/images are
 * preserved. For manual Excel review + openpyxl cross-validation. */
import { mkdirSync, readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { Workbook } from "../dist/index.js";

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "..");
mkdirSync(join(root, "out"), { recursive: true });

// borrow the sample PNG from the image fixture
const imageXlsx = await JSZip.loadAsync(readFileSync(join(root, "tests", "assets", "image.xlsx")));
const png = await imageXlsx.file("xl/media/image1.png").async("uint8array");

const wb = await Workbook.open(join(root, "tests", "assets", "basic.xlsx"));
wb.get("Data").addImage(png, "E5");
wb.coreProperties.title = "M6 Gate Workbook";
wb.coreProperties.author = "FlowDot";
wb.coreProperties.keywords = "charts, images";

const dest = join(root, "out", "m6-media.xlsx");
await wb.save(dest);
console.log("wrote", dest);
