import {
  OpenXmlPackage,
  createRelationshipsDocument,
  nextRelationshipId,
  parseXml,
  relationshipPartPath,
  resolveTargetPath,
  serializeXml,
} from "./openXmlPackage";
import { z } from "zod";

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ELEMENT_NODE = 1;
const EMUS_PER_POINT = 12700;
const CONTENT_TYPE_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const RELATIONSHIP_TYPE_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const DEFAULT_CHART_LEFT = 60;
const DEFAULT_CHART_TOP = 80;
const DEFAULT_CHART_WIDTH = 480;
const DEFAULT_CHART_HEIGHT = 288;

export const slideChartTypeSchema = z.enum(["column", "bar", "line", "pie", "doughnut", "area", "scatter"]);

export const slideChartSeriesSchema = z.object({
  name: z.string().min(1),
  values: z.array(z.number().finite()).min(1),
});

export const slideChartDefinitionSchema = z.object({
  chartType: slideChartTypeSchema,
  title: z.string().optional(),
  categories: z.array(z.string()).optional(),
  series: z.array(slideChartSeriesSchema).min(1),
  stacked: z.boolean().optional(),
  left: z.number().finite().optional(),
  top: z.number().finite().optional(),
  width: z.number().finite().positive().optional(),
  height: z.number().finite().positive().optional(),
}).superRefine((definition, context) => {
  const seriesLength = definition.series[0]?.values.length ?? 0;
  definition.series.forEach((series, index) => {
    if (series.values.length !== seriesLength) {
      context.addIssue({
        code: z.ZodIssueCode.custom,
        message: `series[${index}] values length ${series.values.length} does not match the first series length ${seriesLength}.`,
        path: ["series", index, "values"],
      });
    }
  });
  if (definition.categories && definition.categories.length !== seriesLength) {
    context.addIssue({
      code: z.ZodIssueCode.custom,
      message: `categories length ${definition.categories.length} must match series value length ${seriesLength}.`,
      path: ["categories"],
    });
  }
  if (definition.stacked && ["pie", "doughnut", "scatter"].includes(definition.chartType)) {
    context.addIssue({
      code: z.ZodIssueCode.custom,
      message: `stacked is not supported for ${definition.chartType} charts.`,
      path: ["stacked"],
    });
  }
  if (definition.chartType === "pie" && definition.series.length !== 1) {
    context.addIssue({
      code: z.ZodIssueCode.custom,
      message: "pie charts require exactly one series.",
      path: ["series"],
    });
  }
  if (definition.chartType === "scatter" && definition.categories) {
    definition.categories.forEach((value, index) => {
      if (!Number.isFinite(Number(value))) {
        context.addIssue({
          code: z.ZodIssueCode.custom,
          message: `scatter categories[${index}] must be numeric strings so they can be used as X values.`,
          path: ["categories", index],
        });
      }
    });
  }
});

export type SlideChartType = z.infer<typeof slideChartTypeSchema>;
export type SlideChartSeries = z.infer<typeof slideChartSeriesSchema>;
export type SlideChartDefinition = z.infer<typeof slideChartDefinitionSchema>;

export interface SlideChartMutationResult {
  base64: string;
  xmlShapeId: string;
  chartPartPath: string;
  relationshipId: string;
}

interface ChartTarget {
  slidePath: string;
  slideDoc: XMLDocument;
  slideRelsPath: string;
  slideRelsDoc: XMLDocument;
  frame: Element;
  relationship: Element;
  relationshipId: string;
  chartPartPath: string;
}

interface ChartFrameGeometry {
  left: number;
  top: number;
  width: number;
  height: number;
}

function parseChartDefinition(definition: SlideChartDefinition) {
  return slideChartDefinitionSchema.parse(definition);
}

function escapeXml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function getOnlySlidePath(pkg: OpenXmlPackage) {
  const slidePaths = pkg.listPaths().filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path));
  if (!slidePaths.length) {
    throw new Error("The exported PowerPoint package does not contain a slide XML part.");
  }
  if (slidePaths.length !== 1) {
    throw new Error(`Expected a single-slide PowerPoint export, but found ${slidePaths.length} slide XML parts.`);
  }
  return slidePaths[0];
}

function getOrCreateRelationshipsDoc(pkg: OpenXmlPackage, partPath: string) {
  const relsPath = relationshipPartPath(partPath);
  const relsDoc = pkg.has(relsPath) ? parseXml(pkg.readText(relsPath)) : createRelationshipsDocument();
  return { relsPath, relsDoc };
}

function ensureContentTypeOverride(pkg: OpenXmlPackage, partPath: string, contentType: string) {
  const contentTypesDoc = parseXml(pkg.readText("[Content_Types].xml"));
  const partName = `/${partPath}`;
  const overrides = Array.from(contentTypesDoc.getElementsByTagName("Override"));
  if (!overrides.some((override) => override.getAttribute("PartName") === partName)) {
    const override = contentTypesDoc.createElementNS(contentTypesDoc.documentElement.namespaceURI, "Override");
    override.setAttribute("PartName", partName);
    override.setAttribute("ContentType", contentType);
    contentTypesDoc.documentElement.appendChild(override);
    pkg.writeText("[Content_Types].xml", serializeXml(contentTypesDoc));
  }
}

function removeContentTypeOverride(pkg: OpenXmlPackage, partPath: string) {
  const contentTypesDoc = parseXml(pkg.readText("[Content_Types].xml"));
  const partName = `/${partPath}`;
  Array.from(contentTypesDoc.getElementsByTagName("Override"))
    .filter((override) => override.getAttribute("PartName") === partName)
    .forEach((override) => override.parentNode?.removeChild(override));
  pkg.writeText("[Content_Types].xml", serializeXml(contentTypesDoc));
}

function getSpTree(slideDoc: XMLDocument) {
  const spTree = slideDoc.getElementsByTagNameNS(NS_P, "spTree")[0];
  if (!spTree) {
    throw new Error("The slide XML is missing its shape tree.");
  }
  return spTree;
}

function getAllShapeIds(slideDoc: XMLDocument) {
  return Array.from(slideDoc.getElementsByTagNameNS(NS_P, "cNvPr"))
    .map((node) => Number(node.getAttribute("id") || 0))
    .filter(Number.isFinite);
}

function getNextShapeId(slideDoc: XMLDocument) {
  const ids = getAllShapeIds(slideDoc);
  return String(ids.length ? Math.max(...ids) + 1 : 2);
}

function getNextChartPartPath(pkg: OpenXmlPackage) {
  const chartNumbers = pkg.listPaths()
    .map((path) => /^ppt\/charts\/chart(\d+)\.xml$/.exec(path)?.[1])
    .filter((value): value is string => !!value)
    .map((value) => Number(value));
  const nextNumber = chartNumbers.length ? Math.max(...chartNumbers) + 1 : 1;
  return `ppt/charts/chart${nextNumber}.xml`;
}

function appendRelationship(relationshipsDoc: XMLDocument, type: string, target: string) {
  const relationship = relationshipsDoc.createElementNS(relationshipsDoc.documentElement.namespaceURI, "Relationship");
  const id = nextRelationshipId(relationshipsDoc);
  relationship.setAttribute("Id", id);
  relationship.setAttribute("Type", type);
  relationship.setAttribute("Target", target);
  relationshipsDoc.documentElement.appendChild(relationship);
  return { relationship, relationshipId: id };
}

function getDirectChildByTagName(parent: Element, namespace: string, localName: string) {
  return Array.from(parent.childNodes).find(
    (node) => node.nodeType === ELEMENT_NODE && (node as Element).namespaceURI === namespace && (node as Element).localName === localName,
  ) as Element | undefined;
}

function getChartRelationshipId(frame: Element) {
  const chartElement = frame.getElementsByTagNameNS(NS_C, "chart")[0];
  const relationshipId = chartElement?.getAttributeNS(NS_R, "id") || chartElement?.getAttribute("r:id");
  if (!relationshipId) {
    throw new Error("The chart graphic frame is missing its chart relationship id.");
  }
  return relationshipId;
}

function findChartFrameByXmlShapeId(slideDoc: XMLDocument, xmlShapeId: string) {
  const frame = Array.from(slideDoc.getElementsByTagNameNS(NS_P, "graphicFrame")).find(
    (candidate) => candidate.getElementsByTagNameNS(NS_P, "cNvPr")[0]?.getAttribute("id") === xmlShapeId,
  );
  if (!frame) {
    throw new Error(`Could not find chart frame with XML cNvPr id ${JSON.stringify(xmlShapeId)} on the exported slide.`);
  }
  if (!frame.getElementsByTagNameNS(NS_C, "chart")[0]) {
    throw new Error(`Shape ${JSON.stringify(xmlShapeId)} is not a chart graphic frame.`);
  }
  return frame;
}

function resolveChartTarget(pkg: OpenXmlPackage, xmlShapeId: string): ChartTarget {
  const slidePath = getOnlySlidePath(pkg);
  const slideDoc = parseXml(pkg.readText(slidePath));
  const { relsPath: slideRelsPath, relsDoc: slideRelsDoc } = getOrCreateRelationshipsDoc(pkg, slidePath);
  const frame = findChartFrameByXmlShapeId(slideDoc, xmlShapeId);
  const relationshipId = getChartRelationshipId(frame);
  const relationship = Array.from(slideRelsDoc.getElementsByTagName("Relationship")).find(
    (candidate) => candidate.getAttribute("Id") === relationshipId,
  );
  if (!relationship) {
    throw new Error(`The chart frame ${JSON.stringify(xmlShapeId)} is missing slide relationship ${JSON.stringify(relationshipId)}.`);
  }
  if (relationship.getAttribute("Type") !== RELATIONSHIP_TYPE_CHART) {
    throw new Error(`Shape ${JSON.stringify(xmlShapeId)} does not reference a chart relationship.`);
  }
  return {
    slidePath,
    slideDoc,
    slideRelsPath,
    slideRelsDoc,
    frame,
    relationship,
    relationshipId,
    chartPartPath: resolveTargetPath(slidePath, relationship.getAttribute("Target") || ""),
  };
}

function assertNoUnsupportedChartDependencies(pkg: OpenXmlPackage, chartPartPath: string, action: "update" | "delete") {
  const chartRelsPath = relationshipPartPath(chartPartPath);
  if (!pkg.has(chartRelsPath)) {
    return;
  }
  const chartRelsDoc = parseXml(pkg.readText(chartRelsPath));
  const relationshipCount = chartRelsDoc.getElementsByTagName("Relationship").length;
  if (relationshipCount > 0) {
    throw new Error(`Cannot ${action} chart part ${JSON.stringify(chartPartPath)} because it has dependent chart relationships that this tool does not manage yet.`);
  }
}

function readFrameGeometry(frame: Element): ChartFrameGeometry | null {
  const xfrm = frame.getElementsByTagNameNS(NS_P, "xfrm")[0];
  const off = xfrm?.getElementsByTagNameNS(NS_A, "off")[0];
  const ext = xfrm?.getElementsByTagNameNS(NS_A, "ext")[0];
  const left = Number(off?.getAttribute("x"));
  const top = Number(off?.getAttribute("y"));
  const width = Number(ext?.getAttribute("cx"));
  const height = Number(ext?.getAttribute("cy"));
  if (![left, top, width, height].every(Number.isFinite)) {
    return null;
  }
  return {
    left: left / EMUS_PER_POINT,
    top: top / EMUS_PER_POINT,
    width: width / EMUS_PER_POINT,
    height: height / EMUS_PER_POINT,
  };
}

function toGeometry(definition: SlideChartDefinition, fallback?: ChartFrameGeometry): ChartFrameGeometry {
  return {
    left: definition.left ?? fallback?.left ?? DEFAULT_CHART_LEFT,
    top: definition.top ?? fallback?.top ?? DEFAULT_CHART_TOP,
    width: definition.width ?? fallback?.width ?? DEFAULT_CHART_WIDTH,
    height: definition.height ?? fallback?.height ?? DEFAULT_CHART_HEIGHT,
  };
}

function createElement(doc: XMLDocument, namespace: string, qualifiedName: string, attributes: Record<string, string> = {}) {
  const element = doc.createElementNS(namespace, qualifiedName);
  Object.entries(attributes).forEach(([key, value]) => element.setAttribute(key, value));
  return element;
}

function setFrameTransform(frame: Element, geometry: ChartFrameGeometry) {
  let xfrm = getDirectChildByTagName(frame, NS_P, "xfrm");
  if (!xfrm) {
    xfrm = createElement(frame.ownerDocument, NS_P, "p:xfrm");
    const graphic = getDirectChildByTagName(frame, NS_A, "graphic");
    frame.insertBefore(xfrm, graphic || null);
  }
  while (xfrm.firstChild) {
    xfrm.removeChild(xfrm.firstChild);
  }
  const off = createElement(frame.ownerDocument, NS_A, "a:off", {
    x: String(Math.round(geometry.left * EMUS_PER_POINT)),
    y: String(Math.round(geometry.top * EMUS_PER_POINT)),
  });
  const ext = createElement(frame.ownerDocument, NS_A, "a:ext", {
    cx: String(Math.round(geometry.width * EMUS_PER_POINT)),
    cy: String(Math.round(geometry.height * EMUS_PER_POINT)),
  });
  xfrm.appendChild(off);
  xfrm.appendChild(ext);
}

function createGraphicFrameElement(doc: XMLDocument, xmlShapeId: string, relationshipId: string, geometry: ChartFrameGeometry) {
  const frame = createElement(doc, NS_P, "p:graphicFrame");

  const nvGraphicFramePr = createElement(doc, NS_P, "p:nvGraphicFramePr");
  const cNvPr = createElement(doc, NS_P, "p:cNvPr", { id: xmlShapeId, name: `Chart ${xmlShapeId}` });
  const cNvGraphicFramePr = createElement(doc, NS_P, "p:cNvGraphicFramePr");
  cNvGraphicFramePr.appendChild(createElement(doc, NS_A, "a:graphicFrameLocks", { noGrp: "1" }));
  const nvPr = createElement(doc, NS_P, "p:nvPr");
  nvGraphicFramePr.appendChild(cNvPr);
  nvGraphicFramePr.appendChild(cNvGraphicFramePr);
  nvGraphicFramePr.appendChild(nvPr);

  const xfrm = createElement(doc, NS_P, "p:xfrm");
  xfrm.appendChild(createElement(doc, NS_A, "a:off", {
    x: String(Math.round(geometry.left * EMUS_PER_POINT)),
    y: String(Math.round(geometry.top * EMUS_PER_POINT)),
  }));
  xfrm.appendChild(createElement(doc, NS_A, "a:ext", {
    cx: String(Math.round(geometry.width * EMUS_PER_POINT)),
    cy: String(Math.round(geometry.height * EMUS_PER_POINT)),
  }));

  const graphic = createElement(doc, NS_A, "a:graphic");
  const graphicData = createElement(doc, NS_A, "a:graphicData", {
    uri: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  });
  const chart = createElement(doc, NS_C, "c:chart");
  chart.setAttributeNS(NS_R, "r:id", relationshipId);
  graphicData.appendChild(chart);
  graphic.appendChild(graphicData);

  frame.appendChild(nvGraphicFramePr);
  frame.appendChild(xfrm);
  frame.appendChild(graphic);
  return frame;
}

function insertShapeIntoSpTree(spTree: Element, shape: Element) {
  const extLst = Array.from(spTree.childNodes).find(
    (node) => node.nodeType === ELEMENT_NODE && (node as Element).namespaceURI === NS_P && (node as Element).localName === "extLst",
  );
  spTree.insertBefore(shape, extLst || null);
}

function buildStringPoints(values: string[]) {
  return values.map((value, index) => `<c:pt idx="${index}"><c:v>${escapeXml(value)}</c:v></c:pt>`).join("");
}

function buildNumberPoints(values: number[]) {
  return values.map((value, index) => `<c:pt idx="${index}"><c:v>${Number.isFinite(value) ? value : 0}</c:v></c:pt>`).join("");
}

function buildStringLiteral(values: string[]) {
  return `<c:strLit><c:ptCount val="${values.length}"/>${buildStringPoints(values)}</c:strLit>`;
}

function buildNumberLiteral(values: number[]) {
  return `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${buildNumberPoints(values)}</c:numLit>`;
}

function buildDataLabelsXml() {
  return "<c:dLbls><c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/><c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls>";
}

function buildRichText(text: string) {
  return `<c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(text)}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></c:rich>`;
}

function buildTitleXml(title: string) {
  return `<c:title><c:tx>${buildRichText(title)}</c:tx><c:layout/><c:overlay val="0"/></c:title>`;
}

function buildLegendXml() {
  return "<c:legend><c:legendPos val=\"t\"/><c:layout/><c:overlay val=\"0\"/></c:legend>";
}

function getCategoryValues(definition: SlideChartDefinition) {
  const pointCount = definition.series[0]?.values.length ?? 0;
  return definition.categories || Array.from({ length: pointCount }, (_value, index) => String(index + 1));
}

function getScatterXValues(definition: SlideChartDefinition) {
  const categories = getCategoryValues(definition);
  return categories.map((value, index) => {
    const numeric = Number(value);
    return Number.isFinite(numeric) ? numeric : index + 1;
  });
}

function buildSeriesXml(definition: SlideChartDefinition) {
  const categories = getCategoryValues(definition);
  const scatterXValues = definition.chartType === "scatter" ? getScatterXValues(definition) : null;

  return definition.series.map((series, index) => {
    const header = `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:v>${escapeXml(series.name)}</c:v></c:tx>`;
    const footer = `${buildDataLabelsXml()}</c:ser>`;

    if (definition.chartType === "scatter") {
      return `${header}<c:spPr/><c:xVal>${buildNumberLiteral(scatterXValues || [])}</c:xVal><c:yVal>${buildNumberLiteral(series.values)}</c:yVal>${footer}`;
    }

    const categoryXml = `<c:cat>${buildStringLiteral(categories)}</c:cat>`;
    const valueXml = `<c:val>${buildNumberLiteral(series.values)}</c:val>`;
    const markerXml = definition.chartType === "line" ? "<c:marker><c:symbol val=\"none\"/></c:marker>" : "";
    return `${header}${markerXml}<c:spPr/>${categoryXml}${valueXml}${footer}`;
  }).join("");
}

function buildAxesXml(chartType: SlideChartType) {
  const categoryAxisId = 48650112;
  const valueAxisId = 48672768;

  if (chartType === "scatter") {
    return `
      <c:valAx>
        <c:axId val="${categoryAxisId}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="${valueAxisId}"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="midCat"/>
      </c:valAx>
      <c:valAx>
        <c:axId val="${valueAxisId}"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="out"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="${categoryAxisId}"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="midCat"/>
      </c:valAx>`;
  }

  const categoryAxisPosition = chartType === "bar" ? "l" : "b";
  const valueAxisPosition = chartType === "bar" ? "b" : "l";

  return `
    <c:catAx>
      <c:axId val="${categoryAxisId}"/>
      <c:scaling><c:orientation val="minMax"/></c:scaling>
      <c:delete val="0"/>
      <c:axPos val="${categoryAxisPosition}"/>
      <c:numFmt formatCode="General" sourceLinked="1"/>
      <c:majorTickMark val="out"/>
      <c:minorTickMark val="none"/>
      <c:tickLblPos val="nextTo"/>
      <c:crossAx val="${valueAxisId}"/>
      <c:crosses val="autoZero"/>
      <c:auto val="1"/>
      <c:lblAlgn val="ctr"/>
      <c:lblOffset val="100"/>
      <c:noMultiLvlLbl val="0"/>
    </c:catAx>
    <c:valAx>
      <c:axId val="${valueAxisId}"/>
      <c:scaling><c:orientation val="minMax"/></c:scaling>
      <c:delete val="0"/>
      <c:axPos val="${valueAxisPosition}"/>
      <c:majorGridlines/>
      <c:numFmt formatCode="General" sourceLinked="1"/>
      <c:majorTickMark val="out"/>
      <c:minorTickMark val="none"/>
      <c:tickLblPos val="nextTo"/>
      <c:crossAx val="${categoryAxisId}"/>
      <c:crosses val="autoZero"/>
      <c:crossBetween val="between"/>
    </c:valAx>`;
}

function buildPlotAreaXml(definition: SlideChartDefinition) {
  const seriesXml = buildSeriesXml(definition);
  const grouping = definition.stacked ? "stacked" : "standard";
  const axisIds = "<c:axId val=\"48650112\"/><c:axId val=\"48672768\"/>";

  switch (definition.chartType) {
    case "column":
    case "bar":
      return `<c:plotArea><c:layout/><c:barChart><c:barDir val="${definition.chartType === "column" ? "col" : "bar"}"/><c:grouping val="${definition.stacked ? "stacked" : "clustered"}"/><c:varyColors val="0"/>${seriesXml}${definition.stacked ? "<c:overlap val=\"100\"/>" : ""}${axisIds}</c:barChart>${buildAxesXml(definition.chartType)}</c:plotArea>`;
    case "line":
      return `<c:plotArea><c:layout/><c:lineChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${seriesXml}${axisIds}</c:lineChart>${buildAxesXml(definition.chartType)}</c:plotArea>`;
    case "area":
      return `<c:plotArea><c:layout/><c:areaChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${seriesXml}${axisIds}</c:areaChart>${buildAxesXml(definition.chartType)}</c:plotArea>`;
    case "pie":
      return `<c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/>${seriesXml}<c:firstSliceAng val="0"/></c:pieChart></c:plotArea>`;
    case "doughnut":
      return `<c:plotArea><c:layout/><c:doughnutChart><c:varyColors val="1"/>${seriesXml}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart></c:plotArea>`;
    case "scatter":
      return `<c:plotArea><c:layout/><c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${seriesXml}${axisIds}</c:scatterChart>${buildAxesXml(definition.chartType)}</c:plotArea>`;
  }
}

function buildChartXml(definition: SlideChartDefinition) {
  const effectiveTitle = definition.title?.trim() || "Chart";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">
    <c:style val="2"/>
    <c:lang val="en-US"/>
    <c:roundedCorners val="0"/>
    <c:chart>
      ${buildTitleXml(effectiveTitle)}
      <c:autoTitleDeleted val="0"/>
      ${buildPlotAreaXml(definition)}
      ${buildLegendXml()}
      <c:plotVisOnly val="1"/>
      <c:dispBlanksAs val="gap"/>
      <c:showDLblsOverMax val="0"/>
    </c:chart>
    <c:printSettings>
      <c:headerFooter/>
      <c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>
      <c:pageSetup/>
    </c:printSettings>
  </c:chartSpace>`;
}

export function createChartInBase64Presentation(base64: string, input: SlideChartDefinition): SlideChartMutationResult {
  const definition = parseChartDefinition(input);
  const pkg = new OpenXmlPackage(base64);
  const slidePath = getOnlySlidePath(pkg);
  const slideDoc = parseXml(pkg.readText(slidePath));
  const { relsPath, relsDoc } = getOrCreateRelationshipsDoc(pkg, slidePath);
  const xmlShapeId = getNextShapeId(slideDoc);
  const chartPartPath = getNextChartPartPath(pkg);
  const { relationshipId } = appendRelationship(relsDoc, RELATIONSHIP_TYPE_CHART, `../charts/${chartPartPath.split("/").pop()}`);
  const frame = createGraphicFrameElement(slideDoc, xmlShapeId, relationshipId, toGeometry(definition));

  insertShapeIntoSpTree(getSpTree(slideDoc), frame);
  pkg.writeText(slidePath, serializeXml(slideDoc));
  pkg.writeText(relsPath, serializeXml(relsDoc));
  pkg.writeText(chartPartPath, buildChartXml(definition));
  ensureContentTypeOverride(pkg, chartPartPath, CONTENT_TYPE_CHART);

  return {
    base64: pkg.toBase64(),
    xmlShapeId,
    chartPartPath,
    relationshipId,
  };
}

export function updateChartInBase64Presentation(base64: string, xmlShapeId: string, input: SlideChartDefinition): SlideChartMutationResult {
  const definition = parseChartDefinition(input);
  const pkg = new OpenXmlPackage(base64);
  const target = resolveChartTarget(pkg, xmlShapeId);
  assertNoUnsupportedChartDependencies(pkg, target.chartPartPath, "update");

  setFrameTransform(target.frame, toGeometry(definition, readFrameGeometry(target.frame) || undefined));
  pkg.writeText(target.slidePath, serializeXml(target.slideDoc));
  pkg.writeText(target.chartPartPath, buildChartXml(definition));
  ensureContentTypeOverride(pkg, target.chartPartPath, CONTENT_TYPE_CHART);

  return {
    base64: pkg.toBase64(),
    xmlShapeId,
    chartPartPath: target.chartPartPath,
    relationshipId: target.relationshipId,
  };
}

export function deleteChartInBase64Presentation(base64: string, xmlShapeId: string): SlideChartMutationResult {
  const pkg = new OpenXmlPackage(base64);
  const target = resolveChartTarget(pkg, xmlShapeId);
  assertNoUnsupportedChartDependencies(pkg, target.chartPartPath, "delete");

  target.frame.parentNode?.removeChild(target.frame);
  target.relationship.parentNode?.removeChild(target.relationship);
  pkg.writeText(target.slidePath, serializeXml(target.slideDoc));
  pkg.writeText(target.slideRelsPath, serializeXml(target.slideRelsDoc));
  if (pkg.has(target.chartPartPath)) {
    pkg.delete(target.chartPartPath);
  }
  const chartRelsPath = relationshipPartPath(target.chartPartPath);
  if (pkg.has(chartRelsPath)) {
    pkg.delete(chartRelsPath);
  }
  removeContentTypeOverride(pkg, target.chartPartPath);

  return {
    base64: pkg.toBase64(),
    xmlShapeId,
    chartPartPath: target.chartPartPath,
    relationshipId: target.relationshipId,
  };
}
