"""Author a canonical Excel-style .xlsx (shared strings, cached formula, date
numFmt) by hand — the representation openpyxl does NOT emit but real Excel does.
Exercises the t="s" path, data_only reads, and date-by-numFmt reads.
"""
import os
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "..", "tests", "assets", "shared_strings.xlsx")

DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

entries = {
    "[Content_Types].xml": DECL + (
        f'<Types xmlns="{NS_CT}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    ),
    "_rels/.rels": DECL + (
        f'<Relationships xmlns="{NS_PR}">'
        f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    ),
    "xl/workbook.xml": DECL + (
        f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_R}">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    ),
    "xl/_rels/workbook.xml.rels": DECL + (
        f'<Relationships xmlns="{NS_PR}">'
        f'<Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/>'
        f'<Relationship Id="rId3" Type="{REL}/sharedStrings" Target="sharedStrings.xml"/>'
        '</Relationships>'
    ),
    "xl/sharedStrings.xml": DECL + (
        f'<sst xmlns="{NS_MAIN}" count="4" uniqueCount="3">'
        '<si><t>Hello</t></si>'
        '<si><t>World</t></si>'
        '<si><t xml:space="preserve"> spaced </t></si>'
        '</sst>'
    ),
    "xl/worksheets/sheet1.xml": DECL + (
        f'<worksheet xmlns="{NS_MAIN}">'
        '<dimension ref="A1:B3"/>'
        '<sheetData>'
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'
        '<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>42</v></c></row>'
        '<row r="3"><c r="A3"><f>SUM(B2:B2)</f><v>42</v></c><c r="B3" s="1"><v>46023</v></c></row>'
        '</sheetData>'
        '</worksheet>'
    ),
    "xl/styles.xml": DECL + (
        f'<styleSheet xmlns="{NS_MAIN}">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    ),
}

order = [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/sheet1.xml",
    "xl/sharedStrings.xml",
    "xl/styles.xml",
]

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for name in order:
        z.writestr(name, entries[name])

print("wrote", os.path.normpath(OUT), os.path.getsize(OUT), "bytes")
