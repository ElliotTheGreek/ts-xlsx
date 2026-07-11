/** Chart part (xl/charts/chartN.xml) — read + preserve (authoring deferred). */
import { XmlPart } from "../opc/package.js";
import { ChartInfo, parseChart } from "../oxml/chart.js";

export class ChartPart extends XmlPart {
  get info(): ChartInfo {
    return parseChart(this.root);
  }
}
