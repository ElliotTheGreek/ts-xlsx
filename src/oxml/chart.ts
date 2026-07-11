/**
 * Chart read wrapper — parses a `xl/charts/chartN.xml` (`<chartSpace>` in the
 * DrawingML chart namespace) into a plain, read-only summary: chart type, title,
 * and each series' reference formulas + cached values. Chart *authoring* is a
 * later milestone; existing charts are preserved byte-identical on round-trip.
 */
import { XmlElement } from "../xml/dom.js";
import { nsmap } from "./ns.js";

const C = nsmap.c;
const A = nsmap.a;

/** The plot-type elements ts-xlsx recognizes (localName under plotArea). */
const PLOT_TYPES = [
  "barChart",
  "bar3DChart",
  "lineChart",
  "line3DChart",
  "pieChart",
  "pie3DChart",
  "doughnutChart",
  "areaChart",
  "area3DChart",
  "scatterChart",
  "radarChart",
  "stockChart",
  "surfaceChart",
  "bubbleChart",
  "ofPieChart",
] as const;

export interface SeriesInfo {
  name?: string;
  categoriesRef?: string;
  valuesRef?: string;
  cachedValues: number[];
}

export interface ChartInfo {
  type: string;
  title?: string;
  series: SeriesInfo[];
}

function allText(el: XmlElement, nsUri: string, localName: string): string {
  let out = "";
  for (const t of el.findAllDeep(nsUri, localName)) out += t.text;
  return out;
}

function refFormula(container: XmlElement | null): string | undefined {
  if (container === null) return undefined;
  // numRef/strRef/multiLvlStrRef → <f>
  for (const kind of ["numRef", "strRef", "multiLvlStrRef"]) {
    const ref = container.find(C, kind);
    if (ref !== null) return ref.find(C, "f")?.text ?? undefined;
  }
  return container.find(C, "f")?.text ?? undefined;
}

function cachedValues(valEl: XmlElement | null): number[] {
  if (valEl === null) return [];
  const cache = valEl.find(C, "numRef")?.find(C, "numCache");
  if (cache === undefined || cache === null) return [];
  const out: number[] = [];
  for (const pt of cache.findAll(C, "pt")) {
    const v = pt.find(C, "v")?.text;
    if (v !== undefined) out.push(Number(v));
  }
  return out;
}

/** Parse a chart part's root `<chartSpace>` element into a ChartInfo. */
export function parseChart(root: XmlElement): ChartInfo {
  const chart = root.find(C, "chart");
  const plotArea = chart?.find(C, "plotArea") ?? null;

  let type = "unknown";
  let plot: XmlElement | null = null;
  if (plotArea !== null) {
    for (const t of PLOT_TYPES) {
      const el = plotArea.find(C, t);
      if (el !== null) {
        type = t;
        plot = el;
        break;
      }
    }
  }

  const titleEl = chart?.find(C, "title") ?? null;
  const title = titleEl === null ? undefined : allText(titleEl, A, "t") || undefined;

  const series: SeriesInfo[] = [];
  if (plot !== null) {
    for (const ser of plot.findAll(C, "ser")) {
      series.push({
        name: refFormula(ser.find(C, "tx")) ?? ser.find(C, "tx")?.find(C, "v")?.text ?? undefined,
        categoriesRef: refFormula(ser.find(C, "cat")),
        valuesRef: refFormula(ser.find(C, "val")),
        cachedValues: cachedValues(ser.find(C, "val")),
      });
    }
  }

  return { type, title, series };
}
