import { getSlideById, getSlideByIndex } from "./powerpointNativeContent";
import { loadShapeSummaries, type PowerPointShapeSummary } from "./powerpointShared";
import {
  inspectSlideXmlFromOfficeSlide,
  resolveSlideXmlShapeTarget,
  type ResolvedSlideXmlShapeTarget,
  type SlideXmlInspection,
  type SlideXmlShapeSummary,
} from "./powerpointSlideXml";

const SHAPE_REF_PATTERN = /^slide-id:([^/]+)\/shape:(\d+)$/;
const POSITION_TOLERANCE = 0.5;

export interface PowerPointShapeRefParts {
  slideId: string;
  xmlShapeId: string;
  ref: string;
}

export interface PowerPointShapeSummaryWithRef extends PowerPointShapeSummary {
  slideId: string;
  xmlShapeId: string;
  ref: string;
}

export interface ResolvedPowerPointShapeRefTarget extends ResolvedSlideXmlShapeTarget {
  slide: PowerPoint.Slide;
  slideIndex: number;
}

function assertSlideId(slideId: string) {
  if (!slideId.trim()) {
    throw new Error(`Invalid PowerPoint shape ref slideId ${JSON.stringify(slideId)}.`);
  }
}

function assertXmlShapeId(xmlShapeId: string) {
  if (!/^\d+$/.test(xmlShapeId)) {
    throw new Error(`Invalid PowerPoint shape ref xmlShapeId ${JSON.stringify(xmlShapeId)}.`);
  }
}

function invalidShapeRefError(ref: string) {
  return new Error(`Invalid PowerPoint shape ref ${JSON.stringify(ref)}. Expected format slide-id:<slideId>/shape:<xmlShapeId>.`);
}

function isNearlyEqual(left: number, right: number) {
  return Math.abs(left - right) <= POSITION_TOLERANCE;
}

function inferExpectedXmlShapeType(shapeSummary: PowerPointShapeSummary) {
  const normalizedType = shapeSummary.type.toLowerCase();
  if (shapeSummary.tableInfo || normalizedType.includes("chart") || normalizedType.includes("table")) {
    return "graphicFrame";
  }
  if (normalizedType.includes("picture") || normalizedType.includes("image") || normalizedType.includes("photo")) {
    return "pic";
  }
  if (normalizedType.includes("connector")) {
    return "cxnSp";
  }
  if (normalizedType.includes("group")) {
    return "grpSp";
  }
  if (shapeSummary.text !== undefined || normalizedType.includes("text") || normalizedType.includes("placeholder")) {
    return "sp";
  }
  return null;
}

function normalizeOfficeShapeNameForAlignment(name: string) {
  return /^\(unnamed \d+\)$/.test(name) ? "" : name;
}

function assertOfficeAndXmlShapeAlignment(
  slideId: string,
  shapeSummaries: PowerPointShapeSummary[],
  xmlShapes: SlideXmlShapeSummary[],
  options: { enforceTextPresence?: boolean } = {},
) {
  const { enforceTextPresence = true } = options;
  shapeSummaries.forEach((shapeSummary, index) => {
    const xmlShape = xmlShapes[index];
    if (!xmlShape) return;
    let invariantCount = 0;

    if (xmlShape.box) {
      invariantCount += 1;
      const isAligned = isNearlyEqual(shapeSummary.left, xmlShape.box.left)
        && isNearlyEqual(shapeSummary.top, xmlShape.box.top)
        && isNearlyEqual(shapeSummary.width, xmlShape.box.width)
        && isNearlyEqual(shapeSummary.height, xmlShape.box.height);
      if (!isAligned) {
        throw new Error(
          `Shape order mismatch between Office and slide XML at index ${index} for slide ${JSON.stringify(slideId)}. `
          + `Office box=(${shapeSummary.left}, ${shapeSummary.top}, ${shapeSummary.width}, ${shapeSummary.height}), `
          + `XML box=(${xmlShape.box.left}, ${xmlShape.box.top}, ${xmlShape.box.width}, ${xmlShape.box.height}).`,
        );
      }
    }

    const expectedXmlType = inferExpectedXmlShapeType(shapeSummary);
    if (expectedXmlType) {
      invariantCount += 1;
      if (xmlShape.type !== expectedXmlType) {
        throw new Error(
          `Shape order mismatch between Office and slide XML at index ${index} for slide ${JSON.stringify(slideId)}. `
          + `Office type ${JSON.stringify(shapeSummary.type)} does not align with XML type ${JSON.stringify(xmlShape.type)}.`,
        );
      }
    }

    const normalizedOfficeName = normalizeOfficeShapeNameForAlignment(shapeSummary.name || "");
    if (normalizedOfficeName || xmlShape.name) {
      invariantCount += 1;
      if (normalizedOfficeName !== (xmlShape.name || "")) {
        throw new Error(
          `Shape order mismatch between Office and slide XML at index ${index} for slide ${JSON.stringify(slideId)}. `
          + `Office name ${JSON.stringify(shapeSummary.name)} does not align with XML name ${JSON.stringify(xmlShape.name)}.`,
        );
      }
    }

    if (enforceTextPresence && (shapeSummary.text !== undefined || !!xmlShape.textBody)) {
      invariantCount += 1;
      if ((shapeSummary.text !== undefined) !== !!xmlShape.textBody) {
        throw new Error(
          `Shape order mismatch between Office and slide XML at index ${index} for slide ${JSON.stringify(slideId)}. `
          + `Office text presence does not align with XML text-body presence.`,
        );
      }
    }

    if (invariantCount === 0) {
      throw new Error(
        `Could not verify Office/XML shape alignment at index ${index} for slide ${JSON.stringify(slideId)}. `
        + "The shape lacks stable comparison data for safe ref assignment.",
      );
    }
  });
}

function buildRefParts(slideId: string, xmlShapeId: string): PowerPointShapeRefParts {
  assertSlideId(slideId);
  assertXmlShapeId(xmlShapeId);
  const ref = `slide-id:${encodeURIComponent(slideId)}/shape:${xmlShapeId}`;
  return { slideId, xmlShapeId, ref };
}

export function buildPowerPointShapeRef(slideId: string, xmlShapeId: string) {
  return buildRefParts(slideId, xmlShapeId).ref;
}

export function parsePowerPointShapeRef(ref: string): PowerPointShapeRefParts {
  const match = SHAPE_REF_PATTERN.exec(ref.trim());
  if (!match) {
    throw invalidShapeRefError(ref);
  }

  let slideId: string;
  try {
    slideId = decodeURIComponent(match[1]);
  } catch {
    throw invalidShapeRefError(ref);
  }
  const xmlShapeId = match[2];
  return buildRefParts(slideId, xmlShapeId);
}

export function enrichShapeSummariesWithRefs(
  slideId: string,
  shapeSummaries: PowerPointShapeSummary[],
  xmlShapes: SlideXmlShapeSummary[],
  options: { enforceTextPresence?: boolean } = {},
): PowerPointShapeSummaryWithRef[] {
  if (shapeSummaries.length !== xmlShapes.length) {
    throw new Error(`Shape count mismatch between Office (${shapeSummaries.length}) and slide XML (${xmlShapes.length}) for slide ${JSON.stringify(slideId)}.`);
  }
  assertOfficeAndXmlShapeAlignment(slideId, shapeSummaries, xmlShapes, options);

  return shapeSummaries.map((shapeSummary, index) => ({
    ...shapeSummary,
    slideId,
    xmlShapeId: xmlShapes[index].xmlShapeId,
    ref: buildPowerPointShapeRef(slideId, xmlShapes[index].xmlShapeId),
  }));
}

export async function loadShapeSummariesWithRefs(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  options: Parameters<typeof loadShapeSummaries>[2] = {},
): Promise<PowerPointShapeSummaryWithRef[]> {
  slide.shapes.load("items");
  await context.sync();
  const summaries = await loadShapeSummaries(context, slide.shapes.items, options);
  const inspection = await inspectSlideXmlFromOfficeSlide(context, slide);
  if (!inspection.slideId) {
    throw new Error("Slide XML inspection did not include a stable slide id.");
  }
  return enrichShapeSummariesWithRefs(inspection.slideId, summaries, inspection.shapes, {
    enforceTextPresence: options.includeText !== false,
  });
}

export async function loadSlideShapeSummariesWithRefs(
  context: PowerPoint.RequestContext,
  slideIndex: number,
  options: Parameters<typeof loadShapeSummaries>[2] = {},
) {
  const slide = await getSlideByIndex(context, slideIndex);
  return loadShapeSummariesWithRefs(context, slide, options);
}

export function resolvePowerPointShapeRefInInspection(inspection: SlideXmlInspection, ref: string) {
  return resolveSlideXmlShapeTarget(inspection, parsePowerPointShapeRef(ref));
}

export async function resolvePowerPointShapeRefTarget(
  context: PowerPoint.RequestContext,
  ref: string,
): Promise<ResolvedPowerPointShapeRefTarget> {
  const parsedRef = parsePowerPointShapeRef(ref);
  const { slide, slideIndex } = await getSlideById(context, parsedRef.slideId);
  const inspection = await inspectSlideXmlFromOfficeSlide(context, slide);
  const resolved = resolveSlideXmlShapeTarget(inspection, parsedRef);
  return {
    ...resolved,
    slide,
    slideIndex,
  };
}
