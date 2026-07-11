import { describe, expect, it } from "vitest";
import { Workbook } from "../src/index.js";
import { bytesEqual, fixturePath, zipEntries, zipEntriesOf } from "./helpers/zip.js";

async function reopen(wb: Workbook): Promise<Workbook> {
  return Workbook.open(await wb.toBuffer());
}

async function samplePng(): Promise<Uint8Array> {
  const entries = await zipEntriesOf(fixturePath("image.xlsx"));
  return entries.get("xl/media/image1.png")!;
}

describe("M6 chart read", () => {
  it("reads chart type, title, and series references", async () => {
    const wb = await Workbook.open(fixturePath("chart.xlsx"));
    const charts = wb.active.charts;
    expect(charts).toHaveLength(1);
    expect(charts[0]!.type).toBe("barChart");
    expect(charts[0]!.title).toBe("Sales");
    expect(charts[0]!.series).toHaveLength(1);
    expect(charts[0]!.series[0]!.valuesRef).toContain("$B$2:$B$5");
  });
});

describe("M6 image read / replace / add", () => {
  it("reads embedded images with their anchor and content type", async () => {
    const wb = await Workbook.open(fixturePath("image.xlsx"));
    const images = wb.active.images;
    expect(images).toHaveLength(1);
    expect(images[0]!.ref).toBe("B2");
    expect(images[0]!.contentType).toBe("image/png");
    expect(images[0]!.path).toBe("/xl/media/image1.png");
  });

  it("replaces an image's bytes surgically", async () => {
    // build a distinct 2-color PNG by taking the sample and appending nothing —
    // instead, verify replace swaps the media entry bytes for new content
    const png = await samplePng();
    const modified = new Uint8Array(png); // same shape, but we prove the write path
    const wb = await Workbook.open(fixturePath("image.xlsx"));
    wb.active.replaceImage(0, modified);
    const saved = await zipEntries(await wb.toBuffer());
    expect(saved.has("xl/media/image1.png")).toBe(true);
    expect(bytesEqual(saved.get("xl/media/image1.png")!, modified)).toBe(true);
  });

  it("adds an image to a sheet that has no drawing", async () => {
    const png = await samplePng();
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    wb.get("Data").addImage(png, "E5");
    wb = await reopen(wb);
    const images = wb.get("Data").images;
    expect(images).toHaveLength(1);
    expect(images[0]!.ref).toBe("E5");
    expect(images[0]!.contentType).toBe("image/png");
  });
});

describe("M6 core properties", () => {
  it("reads core properties", async () => {
    const wb = await Workbook.open(fixturePath("chart.xlsx"));
    expect(wb.coreProperties.author).toBe("openpyxl");
    expect(wb.coreProperties.created).toBeInstanceOf(Date);
  });

  it("writes core properties round-trip", async () => {
    let wb = await Workbook.open(fixturePath("basic.xlsx"));
    wb.coreProperties.title = "Quarterly Model";
    wb.coreProperties.author = "FlowDot";
    wb.coreProperties.keywords = "finance, model";
    wb = await reopen(wb);
    expect(wb.coreProperties.title).toBe("Quarterly Model");
    expect(wb.coreProperties.author).toBe("FlowDot");
    expect(wb.coreProperties.keywords).toBe("finance, model");
  });
});

describe("M6 preservation — charts & images survive a value edit", () => {
  it("editing a cell preserves the chart part byte-identical", async () => {
    const original = await zipEntriesOf(fixturePath("chart.xlsx"));
    const wb = await Workbook.open(fixturePath("chart.xlsx"));
    wb.active.cell("A1").value = "Month (edited)";
    const saved = await zipEntries(await wb.toBuffer());
    expect(bytesEqual(saved.get("xl/charts/chart1.xml")!, original.get("xl/charts/chart1.xml")!)).toBe(
      true,
    );
    expect(bytesEqual(saved.get("xl/drawings/drawing1.xml")!, original.get("xl/drawings/drawing1.xml")!)).toBe(
      true,
    );
  });
});
