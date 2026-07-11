# ts-xlsx — Developer Guide

> **Published on npm as [`ts-xlsx-edit`](https://www.npmjs.com/package/ts-xlsx-edit)**
> (`npm install ts-xlsx-edit`). The bare `ts-xlsx` name was already taken; the repo
> and local directory stay `ts-xlsx`.

A pure-TypeScript port of **openpyxl**: open an existing `.xlsx`, read/edit cells,
styles, formulas, sheets, merges, conditional formatting, data validation,
hyperlinks, defined names, images and core properties, and save it back
**preserving every untouched part byte-for-byte**. `jszip` is the only runtime
dependency. Node ≥ 18, ESM, no native modules, no DOM.

This guide is for anyone using the library beyond the happy path or contributing to
it. It explains the architecture, the fidelity guarantees that shape every design
decision, the SpreadsheetML details you need to extend it safely, and how the test
suite proves correctness. Quick usage is at the end; the exported symbols in
`src/index.ts` and their doc-comments are the full API reference.

---

## 1. Why this library exists

The JS ecosystem can create spreadsheets, but faithfully **editing an existing one**
is a gap. SheetJS Community Edition drops cell styling on write (fills/fonts/borders/
conditional formats disappear); `exceljs` creates and reads but is not a
fidelity-preserving round-trip editor. ts-xlsx opens a real workbook, lets you change
what you mean to change, and rewrites it so that **everything you didn't touch is
byte-identical** — stricter even than openpyxl, which drops charts, images, pivot
tables, and stray files on its own round-trip.

---

## 2. Architecture — the layer cake

ts-xlsx is layered. The bottom two layers are generic OOXML infrastructure; only the
layers above them are specific to SpreadsheetML.

```
public API      src/{workbook,worksheet,cell}.ts  + src/styles/values.ts
                src/{util,numberFormats,numberFormat,datetimes}.ts
parts           src/parts/*.ts        — one class per OPC part type
oxml            src/oxml/*.ts         — typed wrappers over SpreadsheetML elements
opc             src/opc/*.ts          — OPC package: parts, rels, content-types, zip
xml             src/xml/*.ts          — lossless XML DOM (parse ⇄ serialize is byte-exact)
```

- **`src/xml/`** — a hand-rolled lossless DOM. Its defining property is
  `serialize(parse(bytes)) === bytes` for any well-formed document: attribute order,
  quote characters, whitespace, self-closing vs. open/close tag shape, the XML
  declaration and a UTF-8 BOM flag are all preserved. Everything above depends on this.
- **`src/opc/`** — the Open Packaging Conventions layer: the zip container, `Part`
  objects, the relationship graph, and `[Content_Types].xml`. It is deliberately
  format-agnostic; `constants.ts`/`spec.ts` supply the SpreadsheetML content/relationship
  types.
- **`src/oxml/`** — typed wrappers (`OxmlWrapper` subclasses) over the raw XML elements:
  `<workbook>`, `<worksheet>`/`<sheetData>`/`<row>`/`<c>`, `<sst>`, `<styleSheet>`,
  `<chartSpace>`. Reads use `find`/`findAll` by namespace URI; writes use `createElement`
  and respect each element's schema child-order.
- **public API** — `Workbook`, `Worksheet`, `Cell`, the style value objects, and the
  coordinate/date/number-format helpers. This is what consumers import.

### Part-type registry

`src/index.ts` maps each content type to its `Part` subclass via
`PartFactory.partTypeFor` (the same idea as openpyxl's reader dispatch). **A new part
class must be registered there**, or the loader falls back to a generic `XmlPart`/`Part`
and your typed accessors never run.

---

## 3. The fidelity model (the invariant everything serves)

Three rules, enforced mechanically rather than by convention:

1. **Open → save with no edits is loss-free.** In `OpcPackage` (`src/opc/package.ts`):
   every part keeps its original bytes; an `XmlPart` re-serializes **only if its document
   was marked dirty**, otherwise it returns the original blob untouched.
   `[Content_Types].xml` and each `.rels` are written back verbatim unless the package
   structure changed. An **opaque sweep** carries through every zip entry not reachable
   from the relationship graph — vendor XML, stray files, anything — which openpyxl drops.
   Output entry order follows the original archive.

2. **An edit touches only what it must.** Dirty-tracking is per-document, so writing one
   cell re-serializes only that worksheet part (plus the shared-strings or styles part if
   an edit genuinely reached it). The test suite asserts the exact set of changed entries.

3. **Reads never mutate.** This is the deliberate divergence from openpyxl's
   mutate-on-access model, and it is *required* by rule 1. A `Cell` is a **lazy proxy**
   over `(worksheet, coordinate)`: reading a value or a style resolves the backing `<c>`
   if it exists and returns `null`/defaults otherwise — it never creates XML. The element
   is materialized only when you assign to the cell.

The litmus test for any new feature: **does an unrelated open→save still produce
byte-identical output?** Add an assertion to that effect — the `mN` test suites show the
pattern.

---

## 4. SpreadsheetML specifics you must know

- **Default namespace.** Unlike PowerPoint (`p:`) and Word (`w:`), worksheet/workbook
  elements live in the **default** namespace (`<worksheet xmlns=".../spreadsheetml/2006/main">`,
  so `<row>`/`<c>` are unprefixed). The namespace registry (`src/oxml/ns.ts`) maps the
  canonical prefix `main` to that URI; `createElement` resolves an in-scope default
  declaration and returns an **unprefixed** element — exactly what editing an existing
  sheet needs. Matching is always by **namespace URI + local name**, never prefix string.
- **`createElement` attach-as-you-go discipline (a real trap).** `createElement(tag,
  context)` resolves the namespace from the context element's *in-scope* declarations. If
  you build a detached subtree — creating children before the parent is attached to the
  tree that declares the default namespace — the children come out wrongly prefixed
  (`main:c`, `xdr:pic`). **Always append the parent into the live tree first, then create
  its children.**
- **Cell value typing** (`@t` on `<c>`): `n` (number, default — attribute omitted) · `s`
  (shared-string index) · `inlineStr` (`<is><t>`) · `str` (formula string result) · `b`
  (boolean `<v>0|1`) · `e` (error string) · `d` (ISO date, rare). A formula cell also
  carries `<f>`.
- **Shared strings.** Real Excel stores text via a shared-string table (`t="s"`, an index
  into `sharedStrings.xml`); some writers (openpyxl included) emit inline strings and no
  SST. ts-xlsx reads both and, on write, **interns into the shared-string table**
  (Excel-canonical), creating the SST part if the workbook lacks one.
- **Style indirection.** A cell's `@s` indexes `styles.xml`'s `<cellXfs>`; each `<xf>`
  references `fontId`/`fillId`/`borderId`/`numFmtId` into the corresponding tables.
  Restyling a cell means **find-or-add a font/fill/border, then find-or-add an `<xf>`**
  equal to the base xf with that id — never mutate a shared xf, which would restyle every
  cell that references it. This find-or-reuse engine (`src/oxml/styles.ts`) keeps the
  tables from bloating; a test asserts three identically-styled cells add exactly one xf.
- **Serial dates.** Excel stores dates as serial numbers under a number format and a
  workbook epoch (1900, including the historical leap-year bug, or 1904). `src/datetimes.ts`
  ports openpyxl's `from_excel`/`to_excel`; `src/numberFormats.ts` ports the built-in
  formats + the date-format heuristic. All date math is done in **UTC** to avoid timezone
  drift on round-trip.
- **Formulas are stored, not evaluated** — a formula engine is out of scope. Writing a
  formula sets `calcPr/@fullCalcOnLoad` and drops the now-stale `calcChain` part so the
  spreadsheet app recomputes on open.
- **Absolute vs. relative relationship targets.** Excel authors `.rels` with **absolute**
  targets (`Target="/xl/worksheets/sheet1.xml"`) where some producers use relative ones.
  The OPC layer resolves both.

---

## 5. Extending the library

1. **Model the XML in an oxml wrapper.** Add read accessors + write helpers to the relevant
   `src/oxml/*.ts` class. Use `find`/`findAll` for reads and `createElement`
   (attach-as-you-go, §4) for writes, inserting new children in schema order via the
   `*_CHILDREN` sequence arrays / `getOrAddOrdered` helpers.
2. **If it's a new part type**, add `src/parts/Xxx.ts` (`XmlPart` for XML, `Part` for
   binary), a `static createNew(pkg)` if it can be authored, and **register its content
   type in `src/index.ts`**.
3. **Expose it on the public class** (`Workbook`/`Worksheet`/`Cell`). Reads return plain
   data and never mutate; route table/xf reuse through `Workbook` so the styles engine
   stays the single source of truth.
4. **Test three ways:** (a) round-trip through the library (`edit → toBuffer → reopen →
   assert`); (b) a fidelity assertion (the edit changes only the expected entries; an
   unrelated open→save stays byte-identical); (c) cross-validate against a reference
   implementation (§6).

---

## 6. Testing philosophy — openpyxl as the oracle

Correctness is not asserted by inspection; it is cross-checked against **openpyxl**, the
mature Python library this is a port of. Install it with `pip install openpyxl`.

openpyxl plays two roles:

- **Spec reference** — its number-format table, date conversion, and element shapes are
  the ported source of truth (values verified verbatim, not guessed).
- **Round-trip oracle** — each milestone's gate script (`scripts/make-*-gate.mjs`) writes
  an edited workbook to `out/`; opening it in openpyxl must succeed with **no repair** and
  read back the expected values/styles/structure. The loop:

  ```bash
  npm run build
  node scripts/make-m3-gate.mjs                  # writes out/m3-styles.xlsx
  python -c "import openpyxl; wb = openpyxl.load_workbook('out/m3-styles.xlsx'); ..."
  ```

The backbone property test (`serialize(parse(bytes)) === bytes` over every XML entry of
every fixture) and a corpus-wide reads-never-mutate test run in CI on every change. The
fixture corpus (`tests/assets/`) mixes openpyxl-authored workbooks, a hand-authored
canonical-Excel workbook exercising the shared-string path, and an `orphan.xlsx` whose
unreferenced entries prove the preservation guarantee (openpyxl drops them; ts-xlsx keeps
them). `tests/helpers/zip.ts::allXlsxInputs()` auto-discovers everything under
`tests/assets/` and `fixtures/`, so dropping a genuine Excel file there immediately
extends coverage.

---

## 7. Directory map

```
src/
  xml/            lossless DOM
  opc/            OPC package layer (parts, rels, content-types, zip)
  oxml/
    ns.ts         namespace registry (main = default ns)
    base.ts       OxmlWrapper + createElement
    simpletypes.ts  attribute converters
    workbook.ts   <workbook>: sheets, definedNames, calcPr, sheet element ops
    worksheet.ts  <worksheet>/<sheetData>/<row>/<c> + merges/cols/panes/CF/DV/hyperlinks/…
    sharedstrings.ts  <sst>/<si>
    styles.ts     <styleSheet>: numFmts/fonts/fills/borders/cellXfs/dxfs + reuse engine
    chart.ts      chart read parser (<chartSpace> → ChartInfo)
  parts/          one class per OPC part type (register each in src/index.ts)
  styles/values.ts  Font/PatternFill/GradientFill/Border/Side/Alignment/Protection/Color
  image/probe.ts  PNG/JPEG/GIF/BMP/TIFF header probe
  workbook.ts worksheet.ts cell.ts   public API
  util.ts     coords (A1 ⇄ row/col, column letters) + EMU units
  numberFormats.ts  built-in formats + is-date-format
  numberFormat.ts   formatValue() display renderer
  datetimes.ts      serial ⇄ Date
  exc.ts            error hierarchy
  index.ts          public exports + PartFactory registrations
tests/    backbone byte-identity + per-milestone API/fidelity + helpers + assets
scripts/  fixture generators, make-*-gate.mjs, verify-pack.mjs
```

---

## 8. Traps (learned the hard way)

- **`createElement` on a detached subtree** → wrong prefixes. Attach the parent first (§4).
- **Mutating a shared `<xf>`/`<font>`** restyles unrelated cells. Always find-or-add via
  the styles engine.
- **Local time in date math** → off-by-hours on round-trip. Always use UTC (`getUTC*` /
  `Date.UTC`); serials are timezone-agnostic.
- **Double-counting the `%` literal** in the number-format renderer — the `%` is already
  in the format's suffix; the percent flag only scales the value by 100.
- **`sst` `count`/`uniqueCount`** are advisory hints; keep them consistent but never rely
  on them for correctness.

---

## 9. Scripts

```bash
npm run build        # tsc (strict) → dist/
npm test             # vitest
npm run lint         # eslint (flat config)
npm run verify-pack  # asserts the npm tarball is dist-only (no source/dev leakage)
```

---

## 10. Not in scope (v1 boundaries, flagged deliberately)

`.xls`/BIFF read/write · a formula **evaluation** engine · **chart authoring** (existing
charts are read + preserved) · **comment authoring** and **named-style authoring**
(existing ones are read + applied) · `copyWorksheet` does not clone the source sheet's own
drawing/hyperlink relationships (cell content + structure are copied).

---

## 11. Usage at a glance

```ts
import { Workbook, Font, PatternFill } from "ts-xlsx-edit";

const wb = await Workbook.open("model.xlsx");   // path | Uint8Array | ArrayBuffer
const ws = wb.active;                            // or wb.get("Sheet1")

ws.cell("B2").value = 42;                         // or ws.cell(2, 2)
ws.cell("C1").value = "=SUM(A1:A10)";             // formula (sets fullCalcOnLoad)
ws.cell("A3").value = new Date(Date.UTC(2030, 0, 1)); // Date → serial + date format
ws.cell("B2").numberFormat = "#,##0.00";
ws.cell("A1").font = new Font({ bold: true, color: "CC0000" });
ws.cell("A1").fill = new PatternFill({ patternType: "solid", fgColor: "FFFF00" });
ws.mergeCells("A5:C5");

for (const row of ws.iterRows({ minRow: 1, maxRow: 10 })) {
  for (const c of row) console.log(c.coordinate, c.value);
}

await wb.save("model-edited.xlsx");               // or const bytes = await wb.toBuffer();
```

Everything except `open`/`save`/`toBuffer` is synchronous. Open with `{ dataOnly: true }`
to read cached formula results instead of formula strings.

Column widths are in character units, row heights in points (as openpyxl exposes them);
A1 and `(row, col)` 1-based addressing are both supported, with `columnIndexFromString` /
`getColumnLetter` / `cellRefToRowCol` / `rowColToCellRef` / `rangeBoundaries` exported.

---

## License

MIT. This library is a port of [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl);
openpyxl's MIT copyright and permission notice are reproduced in `LICENSE` under
"Third-party notices."
