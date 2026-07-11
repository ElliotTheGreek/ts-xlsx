# ts-xlsx — Developer Guide

A pure-TypeScript port of **openpyxl**: open an existing `.xlsx`, read/edit cells,
styles, formulas, sheets, merges, conditional formatting, data validation,
hyperlinks, defined names, images and core properties, and save it back
**preserving every untouched part byte-for-byte**. `jszip` is the only runtime
dependency. Node ≥ 18, ESM, no native modules, no DOM.

This document is for people working **on** ts-xlsx (maintenance, extension, and the
`flowdot-documents` engine that wraps it). For the public API surface, the exported
symbols in `src/index.ts` and their doc-comments are the reference; a short usage
example is at the end.

---

## 1. Architecture — the layer cake

ts-xlsx is a layered port. The bottom two layers are shared, near-verbatim, with the
sibling libraries `ts-pptx` and `ts-docx`; only the format-specific layers above them
are hand-written for SpreadsheetML.

```
public API      src/{workbook,worksheet,cell}.ts  + src/styles/values.ts
                src/{util,numberFormats,numberFormat,datetimes}.ts
parts           src/parts/*.ts        — one class per OPC part type
oxml            src/oxml/*.ts         — typed wrappers over SpreadsheetML elements
opc             src/opc/*.ts          — OPC package: parts, rels, content-types, zip
xml             src/xml/*.ts          — lossless XML DOM (parse ⇄ serialize is byte-exact)
```

**What came verbatim from ts-pptx (do not fork casually):**

- `src/xml/{dom,parser,serializer,escape,index}.ts` — byte-identical across all three
  siblings. This is the lossless DOM whose `serialize(parse(bytes)) === bytes` property
  underpins the whole fidelity story.
- `src/opc/{serialized,content-types,oxml}.ts` and `src/oxml/base.ts` — verbatim.
- `src/opc/{package,rels}.ts` — verbatim **except** the error class name (`XlsxError`).

**What is format-specific (hand-written / adapted):**

- `src/opc/constants.ts`, `src/opc/spec.ts` — SpreadsheetML content types + relationship
  types + default-content-type table.
- `src/opc/packuri.ts` — copied, then **patched** (see §6, the absolute-rel-target fix).
- `src/oxml/ns.ts` — the SpreadsheetML namespace registry.
- `src/oxml/{workbook,worksheet,sharedstrings,styles,chart}.ts` — typed element wrappers.
- `src/parts/*` — `WorkbookPart`, `WorksheetPart`, `SharedStringsPart`, `StylesPart`,
  `CorePropertiesPart`, `ImagePart`, `DrawingPart`, `ChartPart`.
- everything in the public API layer.

### Part-type registry

`src/index.ts` wires content-type → part class into `PartFactory.partTypeFor` (mirrors
openpyxl's reader dispatch). **When you add a new part class, register it there**, or
the loader falls back to a generic `XmlPart`/`Part` and your typed accessors won't run.

---

## 2. The fidelity model (how byte-identity actually works)

This is the load-bearing invariant. Three rules, enforced mechanically:

1. **Open → save with no edits is loss-free.** Implemented in `OpcPackage`
   (`src/opc/package.ts`):
   - every `Part` keeps its **original blob**; an `XmlPart` parses lazily on first
     access and re-serializes **only if its document was marked dirty**. Untouched parts
     return `originalBlob` unchanged.
   - `[Content_Types].xml` and each `.rels` item are written back **verbatim** unless the
     package structure changed (`markStructureDirty()`) or a relationship was added/
     removed. Rels regenerate in deterministic numeric-rId order only when dirty.
   - an **opaque sweep** collects every zip entry not reachable from the rels graph and
     carries it through verbatim. openpyxl **drops** these (vendor XML, stray files);
     ts-xlsx keeps them. See `tests/assets/orphan.xlsx` + `tests/m7.fidelity.test.ts`.
   - output entry order follows the original archive; new items are appended.

2. **An edit touches only what it must.** Because dirty-tracking is per-document, writing
   one cell re-serializes only that sheet part (plus the SST/styles parts if an edit
   genuinely reached them). The `mN` tests assert the exact set of changed entries.

3. **Reads never mutate.** Divergence from openpyxl's mutate-on-access model, required by
   rule 1. `Cell` is a **lazy proxy** over `(worksheet, coordinate)`: reading a value or a
   style resolves the backing `<c>` element if it exists and returns `null`/defaults if it
   doesn't — it never creates XML. Writing calls `_ensureCellCt` which materializes the
   row/cell. See `src/cell.ts`.

If you add a feature, the litmus test is: **does an unrelated open→save still produce
byte-identical output?** Add an assertion to that effect (the `mN` suites show the
pattern).

---

## 3. SpreadsheetML specifics you must know

- **Default namespace.** Unlike PowerPoint (`p:`) and Word (`w:`), worksheet/workbook
  elements live in the **default** namespace (`<worksheet xmlns=".../spreadsheetml/2006/main">`,
  so `<row>`, `<c>` are unprefixed). `src/oxml/ns.ts` maps the canonical prefix `main` to
  that URI, and `nsdecls("main")` emits `xmlns="…"` (not `xmlns:main`). `createElement`
  resolves an in-scope default declaration and returns an **unprefixed** element — which is
  exactly what editing an existing sheet needs. Matching is always by **namespace URI +
  local name**, never by prefix string.
- **createElement attach-as-you-go discipline (TRAP).** `createElement(tag, context)`
  resolves the namespace from the *context element's in-scope declarations*. If you build a
  detached subtree (create children before the parent is attached to the tree that declares
  the default ns), the children come out **wrongly prefixed** (`main:c`, `xdr:pic`). Always
  **append the parent into the live tree first, then create its children.** See
  `src/oxml/styles.ts::getOrAddDxfFill` and `src/parts/drawing.ts::addPicAnchor` for the
  correct ordering.
- **Cell value typing** (`@t` on `<c>`): `n` (number, default — attribute omitted) · `s`
  (shared-string index into `sharedStrings.xml`) · `inlineStr` (`<is><t>`) · `str`
  (formula string result) · `b` (boolean `<v>0|1`) · `e` (error string) · `d` (ISO date,
  rare). A formula cell additionally carries `<f>`. Read logic: `src/cell.ts` `get value`.
- **openpyxl writes `inlineStr` and no `sharedStrings.xml`; real Excel writes `t="s"`.**
  ts-xlsx must handle both on read, and on write **interns into the shared-string table**
  (Excel-canonical), creating the SST part if the workbook lacks one. See
  `Workbook.internString` and `SharedStringsPart.createNew`.
- **Style indirection.** A cell's `@s` indexes `styles.xml`'s `<cellXfs>`; each `<xf>`
  references `fontId`/`fillId`/`borderId`/`numFmtId` into the `<fonts>`/`<fills>`/
  `<borders>`/`<numFmts>` tables. Editing a cell's font means **find-or-add a font, then
  find-or-add an xf** equal to the base xf with that fontId — never mutate a shared xf (it
  would restyle every cell that shares it). This is the `cellXfs`-reuse engine in
  `src/oxml/styles.ts` (`getOrAddXf` + `getOrAdd{Font,Fill,Border}`); the dedup keeps the
  table from bloating (asserted in `tests/m3.styles.test.ts`).
- **Serial dates.** Excel stores dates as serial numbers governed by a number format and a
  workbook epoch (1900, including the leap-year bug, or 1904). `src/datetimes.ts` ports
  openpyxl's `from_excel`/`to_excel`; `src/numberFormats.ts` ports `BUILTIN_FORMATS` +
  `is_date_format`. All date math is done in **UTC** to avoid timezone drift on round-trip.
- **Formulas are stored, not evaluated** (openpyxl's model — a formula engine is explicitly
  out of scope). Writing a formula sets `calcPr/@fullCalcOnLoad` and **drops the stale
  `calcChain` part** so Excel rebuilds it. See `Workbook.setFullCalcOnLoad` /
  `invalidateCalcChain`.
- **Absolute relationship targets.** Excel authors `.rels` with **absolute** targets
  (`Target="/xl/worksheets/sheet1.xml"`) where PowerPoint/Word use relative ones. This is
  the one patch to the otherwise-verbatim OPC layer (§6).

---

## 4. Directory map

```
src/
  xml/            lossless DOM (verbatim from ts-pptx)
  opc/            OPC package layer (constants/spec/packuri adapted; rest verbatim)
  oxml/           SpreadsheetML element wrappers
    ns.ts         namespace registry (main = default ns)
    base.ts       OxmlWrapper + createElement (verbatim)
    simpletypes.ts  attribute converters (verbatim)
    workbook.ts   <workbook>: sheets, definedNames, calcPr, sheet element ops
    worksheet.ts  <worksheet>/<sheetData>/<row>/<c> + merges/cols/panes/CF/DV/hyperlinks/…
    sharedstrings.ts  <sst>/<si>
    styles.ts     <styleSheet>: numFmts/fonts/fills/borders/cellXfs/dxfs + reuse engine
    chart.ts      chart read parser (<chartSpace> → ChartInfo)
  parts/          one class per OPC part type (register each in src/index.ts)
  styles/values.ts  Font/PatternFill/GradientFill/Border/Side/Alignment/Protection/Color
  image/probe.ts  PNG/JPEG/GIF/BMP/TIFF header probe (verbatim from ts-docx)
  workbook.ts worksheet.ts cell.ts   public API
  util.ts     coords (A1 ⇄ row/col, column letters) + EMU units
  numberFormats.ts  BUILTIN_FORMATS + is_date_format
  numberFormat.ts   formatValue() display renderer
  datetimes.ts      serial ⇄ Date
  exc.ts            XlsxError hierarchy
  index.ts          public exports + PartFactory registrations
tests/
  xml/roundtrip.test.ts  opc/roundtrip.test.ts   backbone byte-identity
  m1..m7 *.test.ts       per-milestone API + fidelity
  helpers/zip.ts         allXlsxInputs(), zipEntries(), bytesEqual(), firstDiff()
  assets/*.xlsx          fixture corpus (see §7)
scripts/
  _gen_fixtures.py            openpyxl-authored fixtures
  _gen_sharedstrings_fixture.py  hand-authored Excel-style t="s" fixture
  make-m1..m6-gate.mjs       produce out/*.xlsx for manual Excel + openpyxl review
  verify-pack.mjs            publish tarball gate (dist-only)
```

---

## 5. How to add a feature

1. **Model the XML in an oxml wrapper.** Add read accessors + write helpers to the relevant
   `src/oxml/*.ts` class, extending `OxmlWrapper`. Use `find(nsUri, localName)` /
   `findAll` for reads and `createElement` (attach-as-you-go, §3) for writes. Respect the
   element's child-order sequence — the `*_CHILDREN` arrays + `getOrAddOrdered` helpers
   insert new children in schema order.
2. **If it's a new part type**, add a `src/parts/Xxx.ts` (extend `XmlPart` for XML, `Part`
   for binary), a `static createNew(pkg)` if it can be authored, and **register its content
   type in `src/index.ts`**.
3. **Expose it on the public class** (`Workbook`/`Worksheet`/`Cell`). Reads go through the
   wrapper and return plain data (never mutate). Writes route table/xf reuse through
   `Workbook` so the styles engine stays the single source of truth.
4. **Test three ways:** (a) round-trip through ts-xlsx (`edit → toBuffer → reopen → assert`);
   (b) a fidelity assertion (the edit changes only the expected entries; unrelated open→save
   stays byte-identical); (c) **cross-validate against openpyxl** (§6).

---

## 6. The openpyxl oracle (how correctness is proven)

openpyxl **3.1.5** is available with a working interpreter at:

```
C:\Users\ellio\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe
```

It is used exactly as ts-docx used python-docx:

- **As a spec reference** — its `styles/numbers.py`, `utils/datetime.py`, etc. are the
  ported source of truth (values verified verbatim, not guessed).
- **As a cross-validation oracle** — every milestone's gate script writes a workbook to
  `out/`; opening it in openpyxl must succeed with **no repair** and read back the expected
  values/styles/structure. The full loop:

```bash
npm run build
node scripts/make-m3-gate.mjs          # writes out/m3-styles.xlsx
python.exe -c "import openpyxl; wb=openpyxl.load_workbook('out/m3-styles.xlsx'); ..."
```

**Important divergence caught this way:** the OPC `PackURI.fromRelRef` (copied verbatim from
ts-pptx) mis-resolved **absolute** rel targets. openpyxl/Excel write `Target="/xl/…"`;
`path.posix.join("/xl", "/xl/…")` produced `/xl/xl/…`, so worksheet parts fell to
opaque-preservation (which masked it in the M0 byte-round-trip) and rId resolution threw at
read time. `src/opc/packuri.ts` now takes a leading-slash target as already
package-absolute. This is the **only** intentional edit to the otherwise-verbatim OPC layer.

---

## 7. Fixture corpus (`tests/assets/`)

- **openpyxl-authored** (`scripts/_gen_fixtures.py`): `basic` (styled cells, merges, widths,
  freeze, multi-sheet, hidden sheet), `formulas`, `types` (dates/bools/unicode/repeated
  strings), `chart` (bar chart), `image` (embedded PNG), `comments`, `rules` (CF + data
  validation + hyperlink + defined name).
- **hand-authored, canonical Excel style** (`scripts/_gen_sharedstrings_fixture.py`):
  `shared_strings.xlsx` — the `t="s"` shared-string path plus a **cached** formula and a
  date `numFmt`, which openpyxl does not emit. This is what exercises the dominant
  real-world representation.
- **orphan.xlsx** — `basic.xlsx` + two entries unreachable from the rels graph
  (`xl/vendorData/custom.xml`, `customNotes.txt`). openpyxl drops both on re-save; ts-xlsx
  must keep them.

`tests/helpers/zip.ts::allXlsxInputs()` auto-discovers everything under `tests/assets/` and
`fixtures/`, so dropping a genuine Excel-authored workbook in there immediately extends the
backbone round-trip + reads-never-mutate coverage.

> **Corpus caveat.** The openpyxl-authored fixtures cannot fully exercise the
> "preserve what openpyxl drops" guarantee for pivot tables / native-Excel vendor XML. The
> chart/image/orphan fixtures cover the differentiator that automation can reach; adding a
> few genuine Excel files (with a pivot table) would strengthen it further.

---

## 8. Traps (learned the hard way)

- **createElement on a detached subtree** → wrong prefixes. Attach parent first (§3).
- **Mutating a shared `<xf>`/`<font>`** restyles unrelated cells. Always find-or-add via the
  styles engine.
- **Absolute rel targets** — don't "fix" `packuri.ts` back to the verbatim ts-pptx version.
- **Local time in date math** — always use UTC `getUTC*` / `Date.UTC` (serials are TZ-agnostic).
- **Double-counting the `%` literal** in the number-format renderer — the `%` is already in
  the format's suffix; the `percent` flag only scales by 100.
- **`sst` count hints** are advisory; keep `count`/`uniqueCount` consistent but don't rely on
  them for correctness.

---

## 9. Build, test, publish

```bash
npm run build        # tsc (strict) → dist/
npm test             # vitest (76 tests)
npm run lint         # eslint (flat config)
npm run verify-pack  # asserts the npm tarball is dist-only (no source leakage)
```

`prepublishOnly` runs clean → lint → test → build → verify-pack. Publishing is a tagged-
release flow (`.github/workflows/publish.yml`, OIDC trusted publisher). **This repo is not
published from here** — tagging `vX.Y.Z` is the maintainer's action.

---

## 10. Deferred (v1 boundaries, flagged deliberately)

`.xls`/BIFF read/write · a formula **evaluation** engine · **chart authoring** (existing
charts are read + preserved) · **comment authoring** and **named-style authoring** (existing
ones are read + applied) · `copyWorksheet` does not clone the source sheet's own drawing/
hyperlink relationships in v1 (cell content + structure are copied).

---

## Appendix — usage at a glance

```ts
import { Workbook, Font, PatternFill } from "ts-xlsx";

const wb = await Workbook.open("model.xlsx");   // path | Uint8Array | ArrayBuffer
const ws = wb.active;                            // or wb.get("Sheet1")
ws.cell("B2").value = 42;                         // or ws.cell(2, 2)
ws.cell("C1").value = "=SUM(A1:A10)";             // formula (sets fullCalcOnLoad)
ws.cell("B2").numberFormat = "#,##0.00";
ws.cell("A1").font = new Font({ bold: true, color: "CC0000" });
ws.cell("A1").fill = new PatternFill({ patternType: "solid", fgColor: "FFFF00" });
ws.mergeCells("A5:C5");
await wb.save("model-edited.xlsx");               // or const bytes = await wb.toBuffer();
```

Everything except `open`/`save`/`toBuffer` is synchronous. MIT © FlowDot LLC; openpyxl's MIT
notice is reproduced in `LICENSE`.
