import { OpenXmlPackage, parseXml, serializeXml } from "./openXmlPackage";
import { getSlideByIndex } from "./powerpointNativeContent";

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const ELEMENT_NODE = 1;
const TEXT_NODE = 3;
const PROCESSING_INSTRUCTION_NODE = 7;
const COMMENT_NODE = 8;
const DOCUMENT_TYPE_NODE = 10;
const EMUS_PER_POINT = 12700;
const SHAPE_LOCAL_NAMES = new Set(["sp", "cxnSp", "pic", "graphicFrame", "grpSp", "contentPart"]);

export interface SlideXmlShapeBox {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface SlideXmlShapeSummary {
  index: number;
  xmlShapeId: string;
  name: string;
  type: string;
  shapeElement: Element;
  textBody: Element | null;
  box: SlideXmlShapeBox | null;
}

export interface SlideXmlInspection {
  slideId?: string;
  slidePath: string;
  slideDoc: XMLDocument;
  shapes: SlideXmlShapeSummary[];
}

export interface SlideXmlShapeTargetIdentity {
  slideId: string;
  xmlShapeId: string;
  ref: string;
}

export interface ResolvedSlideXmlShapeTarget extends SlideXmlShapeTargetIdentity {
  inspection: SlideXmlInspection;
  shapeIndex: number;
  shape: SlideXmlShapeSummary;
}

export interface ShapeParagraphXmlReplacement {
  target: SlideXmlShapeTargetIdentity;
  paragraphsXml: string[];
}

function getOnlySlidePath(pkg: OpenXmlPackage) {
  const slides = pkg.listPaths().filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path));
  if (!slides.length) {
    throw new Error("The exported PowerPoint package does not contain a slide XML part.");
  }
  if (slides.length !== 1) {
    throw new Error(`Expected a single-slide PowerPoint export, but found ${slides.length} slide XML parts. Export one slide or pass an explicit slide target.`);
  }
  return slides.sort((left, right) => left.localeCompare(right, undefined, { numeric: true }))[0];
}

function getDirectChildren(parent: Element, namespace: string, localName: string) {
  return Array.from(parent.childNodes).filter(
    (node) => node.nodeType === ELEMENT_NODE
      && (node as Element).namespaceURI === namespace
      && (node as Element).localName === localName,
  ) as Element[];
}

function getSlideShapeElementsInOrder(slideDoc: XMLDocument) {
  const spTree = slideDoc.getElementsByTagNameNS(NS_P, "spTree")[0];
  if (!spTree) {
    throw new Error("The slide XML is missing its shape tree.");
  }

  return Array.from(spTree.childNodes).filter(
    (node) => node.nodeType === ELEMENT_NODE
      && (node as Element).namespaceURI === NS_P
      && SHAPE_LOCAL_NAMES.has((node as Element).localName),
  ) as Element[];
}

function getShapeNonVisualProperties(shape: Element) {
  const candidate = Array.from(shape.childNodes).find(
    (node) => node.nodeType === ELEMENT_NODE
      && (node as Element).namespaceURI === NS_P
      && /^nv/.test((node as Element).localName),
  ) as Element | undefined;
  if (!candidate) {
    throw new Error(`The exported slide XML is missing non-visual properties for <p:${shape.localName}>.`);
  }
  return candidate;
}

function getXmlShapeId(shape: Element, shapeIndex: number) {
  const cNvPr = getShapeNonVisualProperties(shape).getElementsByTagNameNS(NS_P, "cNvPr")[0];
  const xmlShapeId = cNvPr?.getAttribute("id");
  if (!xmlShapeId) {
    throw new Error(`The exported slide XML is missing a non-visual shape id for shapeIndex ${shapeIndex}.`);
  }
  return xmlShapeId;
}

function getXmlShapeName(shape: Element) {
  return getShapeNonVisualProperties(shape).getElementsByTagNameNS(NS_P, "cNvPr")[0]?.getAttribute("name") || "";
}

function getShapeTextBody(shape: Element) {
  return Array.from(shape.childNodes).find(
    (node) => node.nodeType === ELEMENT_NODE
      && (node as Element).namespaceURI === NS_P
      && (node as Element).localName === "txBody",
  ) as Element | null || null;
}

function getShapeTransform(shape: Element) {
  const directTransform = Array.from(shape.childNodes).find(
    (node) => node.nodeType === ELEMENT_NODE
      && (((node as Element).namespaceURI === NS_A) || ((node as Element).namespaceURI === NS_P))
      && (node as Element).localName === "xfrm",
  ) as Element | undefined;
  if (directTransform) return directTransform;

  const propertyContainer = Array.from(shape.childNodes).find(
    (node) => node.nodeType === ELEMENT_NODE
      && (node as Element).namespaceURI === NS_P
      && ["spPr", "grpSpPr"].includes((node as Element).localName),
  ) as Element | undefined;
  if (!propertyContainer) {
    return null;
  }

  return Array.from(propertyContainer.childNodes).find(
    (node) => node.nodeType === ELEMENT_NODE
      && (((node as Element).namespaceURI === NS_A) || ((node as Element).namespaceURI === NS_P))
      && (node as Element).localName === "xfrm",
  ) as Element | undefined || null;
}

function getShapeBoxInPoints(shape: Element): SlideXmlShapeBox | null {
  const xfrm = getShapeTransform(shape);
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

function getShapeSummary(shape: Element, index: number): SlideXmlShapeSummary {
  return {
    index,
    xmlShapeId: getXmlShapeId(shape, index),
    name: getXmlShapeName(shape),
    type: shape.localName,
    shapeElement: shape,
    textBody: getShapeTextBody(shape),
    box: getShapeBoxInPoints(shape),
  };
}

function getShapeByXmlShapeId(slideDoc: XMLDocument, xmlShapeId: string) {
  const shapeIndex = getSlideShapeElementsInOrder(slideDoc).findIndex(
    (candidate, index) => getXmlShapeId(candidate, index) === xmlShapeId,
  );
  if (shapeIndex < 0) {
    throw new Error(`Could not find shape with XML cNvPr id ${JSON.stringify(xmlShapeId)} on the exported slide.`);
  }

  const shape = getSlideShapeElementsInOrder(slideDoc)[shapeIndex];
  return { shape, shapeIndex };
}

function createEmptyParagraph(doc: XMLDocument) {
  const paragraph = doc.createElementNS(NS_A, "a:p");
  paragraph.appendChild(doc.createElementNS(NS_A, "a:endParaRPr"));
  return paragraph;
}

function describeUnsupportedParagraphNode(node: ChildNode) {
  if (node.nodeType === COMMENT_NODE) return "comments are not allowed";
  if (node.nodeType === PROCESSING_INSTRUCTION_NODE) return "processing instructions are not allowed";
  if (node.nodeType === DOCUMENT_TYPE_NODE) return "DOCTYPE nodes are not allowed";
  return `unsupported XML node type ${node.nodeType}`;
}

function validateParagraphNode(node: ChildNode, xmlShapeId: string, paragraphIndex: number): void {
  if (node.nodeType === COMMENT_NODE || node.nodeType === PROCESSING_INSTRUCTION_NODE || node.nodeType === DOCUMENT_TYPE_NODE) {
    throw new Error(`Invalid paragraph XML for shape ${xmlShapeId} at paragraph ${paragraphIndex}: ${describeUnsupportedParagraphNode(node)}.`);
  }
  if (node.nodeType === TEXT_NODE) {
    return;
  }
  if (node.nodeType === ELEMENT_NODE) {
    Array.from(node.childNodes).forEach((child) => validateParagraphNode(child, xmlShapeId, paragraphIndex));
    return;
  }
  throw new Error(`Invalid paragraph XML for shape ${xmlShapeId} at paragraph ${paragraphIndex}: ${describeUnsupportedParagraphNode(node)}.`);
}

function parseParagraphXml(xmlShapeId: string, paragraphXml: string, paragraphIndex: number) {
  let parsed: XMLDocument;
  try {
    parsed = parseXml(`<root xmlns:a="${NS_A}">${paragraphXml}</root>`);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Invalid paragraph XML for shape ${xmlShapeId} at paragraph ${paragraphIndex}: ${message}`);
  }

  const root = parsed.documentElement;
  const elementChildren = Array.from(root.childNodes).filter((node) => node.nodeType === ELEMENT_NODE) as Element[];
  if (elementChildren.length !== 1 || elementChildren[0].namespaceURI !== NS_A || elementChildren[0].localName !== "p") {
    throw new Error(`Invalid paragraph XML for shape ${xmlShapeId} at paragraph ${paragraphIndex}: expected a single <a:p> root element.`);
  }
  const nonWhitespaceTextNodes = Array.from(root.childNodes).filter(
    (node) => node.nodeType === TEXT_NODE && (node.textContent || "").trim().length > 0,
  );
  if (nonWhitespaceTextNodes.length > 0) {
    throw new Error(`Invalid paragraph XML for shape ${xmlShapeId} at paragraph ${paragraphIndex}: unexpected text outside the root <a:p> element.`);
  }
  Array.from(root.childNodes)
    .filter((node) => node.nodeType !== TEXT_NODE && node !== elementChildren[0])
    .forEach((node) => {
      throw new Error(`Invalid paragraph XML for shape ${xmlShapeId} at paragraph ${paragraphIndex}: ${describeUnsupportedParagraphNode(node)}.`);
    });
  validateParagraphNode(elementChildren[0], xmlShapeId, paragraphIndex);
  return elementChildren[0];
}

function importNode(targetDoc: XMLDocument, node: Element) {
  return typeof targetDoc.importNode === "function"
    ? targetDoc.importNode(node, true)
    : node.cloneNode(true);
}

function assertSlideIdMatch(inspection: SlideXmlInspection, target: SlideXmlShapeTargetIdentity) {
  if (!inspection.slideId) {
    throw new Error("Slide XML inspection does not include a slideId, so public shape refs cannot be resolved.");
  }
  if (inspection.slideId !== target.slideId) {
    throw new Error(`Shape ref ${JSON.stringify(target.ref)} targets slide ${JSON.stringify(target.slideId)}, but the exported slide belongs to ${JSON.stringify(inspection.slideId)}.`);
  }
}

function assertNoDuplicateTargets(replacements: ShapeParagraphXmlReplacement[]) {
  const seen = new Set<string>();
  for (const replacement of replacements) {
    const key = `${replacement.target.slideId}::${replacement.target.xmlShapeId}`;
    if (seen.has(key)) {
      throw new Error(`Duplicate shape target ${JSON.stringify(replacement.target.ref)} in one batch edit is not allowed.`);
    }
    seen.add(key);
  }
}

export function inspectSlideXmlFromBase64Presentation(base64: string, options: { slideId?: string } = {}): SlideXmlInspection {
  const pkg = new OpenXmlPackage(base64);
  const slidePath = getOnlySlidePath(pkg);
  const slideDoc = parseXml(pkg.readText(slidePath));
  return {
    slideId: options.slideId,
    slidePath,
    slideDoc,
    shapes: getSlideShapeElementsInOrder(slideDoc).map(getShapeSummary),
  };
}

export async function inspectSlideXmlFromOfficeSlide(context: PowerPoint.RequestContext, slide: PowerPoint.Slide) {
  slide.load("id");
  const exported = slide.exportAsBase64();
  await context.sync();
  return inspectSlideXmlFromBase64Presentation(exported.value, { slideId: slide.id });
}

export async function inspectSlideXmlByIndex(context: PowerPoint.RequestContext, slideIndex: number) {
  const slide = await getSlideByIndex(context, slideIndex);
  return inspectSlideXmlFromOfficeSlide(context, slide);
}

export function resolveSlideXmlShapeTarget(
  inspection: SlideXmlInspection,
  target: SlideXmlShapeTargetIdentity,
): ResolvedSlideXmlShapeTarget {
  assertSlideIdMatch(inspection, target);
  const shapeIndex = inspection.shapes.findIndex((shape) => shape.xmlShapeId === target.xmlShapeId);
  if (shapeIndex < 0) {
    throw new Error(`Could not find shape ref ${JSON.stringify(target.ref)} on exported slide ${JSON.stringify(target.slideId)}.`);
  }
  return {
    ...target,
    inspection,
    shapeIndex,
    shape: inspection.shapes[shapeIndex],
  };
}

export function getShapeParagraphXml(shape: Element) {
  const textBody = getShapeTextBody(shape);
  if (!textBody) {
    return [];
  }

  return getDirectChildren(textBody, NS_A, "p").map((paragraph) => new XMLSerializer().serializeToString(paragraph));
}

export function getShapeParagraphXmlByXmlShapeId(slideDoc: XMLDocument, xmlShapeId: string) {
  return getShapeParagraphXml(getShapeByXmlShapeId(slideDoc, xmlShapeId).shape);
}

export function getShapeParagraphXmlByTarget(target: ResolvedSlideXmlShapeTarget) {
  return getShapeParagraphXml(target.shape.shapeElement);
}

export function replaceShapeParagraphXml(shape: Element, paragraphsXml: string[]) {
  const textBody = getShapeTextBody(shape);
  if (!textBody) {
    throw new Error(`Shape ${JSON.stringify(getXmlShapeName(shape) || getXmlShapeId(shape, -1))} does not contain a text body.`);
  }

  const replacementParagraphs = paragraphsXml.length
    ? paragraphsXml.map((paragraphXml, paragraphIndex) => parseParagraphXml(getXmlShapeId(shape, -1), paragraphXml, paragraphIndex))
    : [createEmptyParagraph(textBody.ownerDocument)];
  const preservedChildren = Array.from(textBody.childNodes).filter(
    (node) => !(node.nodeType === ELEMENT_NODE && (node as Element).namespaceURI === NS_A && (node as Element).localName === "p"),
  );

  while (textBody.firstChild) {
    textBody.removeChild(textBody.firstChild);
  }

  for (const child of preservedChildren) {
    textBody.appendChild(child);
  }
  for (const paragraph of replacementParagraphs) {
    textBody.appendChild(importNode(textBody.ownerDocument, paragraph));
  }
}

export function replaceShapeParagraphXmlByTarget(target: ResolvedSlideXmlShapeTarget, paragraphsXml: string[]) {
  replaceShapeParagraphXml(target.shape.shapeElement, paragraphsXml);
}

export function replaceShapeParagraphXmlInSlideDocument(slideDoc: XMLDocument, replacements: Array<{ xmlShapeId: string; paragraphsXml: string[] }>) {
  const seen = new Set<string>();
  for (const replacement of replacements) {
    if (seen.has(replacement.xmlShapeId)) {
      throw new Error(`Duplicate shape target ${JSON.stringify(replacement.xmlShapeId)} in one batch edit is not allowed.`);
    }
    seen.add(replacement.xmlShapeId);
  }

  const prepared = replacements.map((replacement) => ({
    replacement,
    shape: getShapeByXmlShapeId(slideDoc, replacement.xmlShapeId).shape,
    paragraphs: replacement.paragraphsXml.length
      ? replacement.paragraphsXml.map((paragraphXml, paragraphIndex) => parseParagraphXml(replacement.xmlShapeId, paragraphXml, paragraphIndex))
      : null,
  }));

  for (const entry of prepared) {
    const textBody = getShapeTextBody(entry.shape);
    if (!textBody) {
      throw new Error(`Shape ${entry.replacement.xmlShapeId} does not contain a text body.`);
    }
  }

  for (const entry of prepared) {
    const textBody = getShapeTextBody(entry.shape)!;
    const preservedChildren = Array.from(textBody.childNodes).filter(
      (node) => !(node.nodeType === ELEMENT_NODE && (node as Element).namespaceURI === NS_A && (node as Element).localName === "p"),
    );
    while (textBody.firstChild) {
      textBody.removeChild(textBody.firstChild);
    }
    for (const child of preservedChildren) {
      textBody.appendChild(child);
    }
    const paragraphs = entry.paragraphs || [createEmptyParagraph(textBody.ownerDocument)];
    for (const paragraph of paragraphs) {
      textBody.appendChild(importNode(textBody.ownerDocument, paragraph));
    }
  }
}

export function replaceShapeParagraphXmlInSlideInspection(inspection: SlideXmlInspection, replacements: ShapeParagraphXmlReplacement[]) {
  assertNoDuplicateTargets(replacements);
  const prepared = replacements.map((replacement) => ({
    resolved: resolveSlideXmlShapeTarget(inspection, replacement.target),
    paragraphs: replacement.paragraphsXml,
  }));

  prepared.forEach((entry) => {
    entry.paragraphs.forEach((paragraphXml, paragraphIndex) => {
      parseParagraphXml(entry.resolved.xmlShapeId, paragraphXml, paragraphIndex);
    });
  });

  for (const entry of prepared) {
    replaceShapeParagraphXmlByTarget(entry.resolved, entry.paragraphs);
  }
}

export function replaceShapeParagraphXmlInBase64Presentation(
  base64: string,
  replacements: ShapeParagraphXmlReplacement[],
  options: { slideId: string },
) {
  const pkg = new OpenXmlPackage(base64);
  const inspection = inspectSlideXmlFromBase64Presentation(base64, { slideId: options.slideId });
  replaceShapeParagraphXmlInSlideInspection(inspection, replacements);
  pkg.writeText(inspection.slidePath, serializeXml(inspection.slideDoc));
  return pkg.toBase64();
}
