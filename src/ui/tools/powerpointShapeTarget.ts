import { findSlideShapeIndexByXmlShapeIdInBase64Presentation } from "./powerpointOpenXml";
import { formatAvailableShapeTargets, isPowerPointRequirementSetSupported } from "./powerpointShared";

export interface ResolvedPowerPointShapeTarget {
  shape: PowerPoint.Shape;
  shapeId: string;
  shapeIndex: number;
}

export interface ResolvedShapeIdentity {
  officeShapeId: string;
  xmlShapeId?: string;
  shapeIndex: number;
}

async function exportSlideShapeIds(context: PowerPoint.RequestContext, slide: PowerPoint.Slide) {
  if (!isPowerPointRequirementSetSupported("1.8")) return null;
  const exported = slide.exportAsBase64();
  await context.sync();
  return exported.value;
}

export async function remapShapeIdentityAfterRoundTrip(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  slideIndex: number,
  requestedShapeId: string | number,
  shapeIdMap?: Record<string, string>,
): Promise<ResolvedShapeIdentity> {
  const requestedId = String(requestedShapeId);

  slide.shapes.load("items/id,name");
  await context.sync();

  const directMatchIndex = slide.shapes.items.findIndex((shape) => shape.id === requestedId);
  if (directMatchIndex >= 0) {
    return {
      officeShapeId: slide.shapes.items[directMatchIndex].id,
      xmlShapeId: requestedId,
      shapeIndex: directMatchIndex,
    };
  }

  const exportedBase64 = await exportSlideShapeIds(context, slide);
  const candidateIds = [requestedId];
  if (shapeIdMap?.[requestedId] && shapeIdMap[requestedId] !== requestedId) {
    candidateIds.unshift(shapeIdMap[requestedId]);
  }

  for (const candidateId of candidateIds) {
    if (!exportedBase64) break;
    const xmlMatchIndex = findSlideShapeIndexByXmlShapeIdInBase64Presentation(exportedBase64, candidateId);
    if (xmlMatchIndex >= 0) {
      const xmlMatch = slide.shapes.items[xmlMatchIndex];
      if (xmlMatch) {
        return {
          officeShapeId: xmlMatch.id,
          xmlShapeId: candidateId,
          shapeIndex: xmlMatchIndex,
        };
      }
    }
  }

  throw new Error(`Shape ${requestedId} was not found on slide ${slideIndex + 1}. ${formatAvailableShapeTargets(slideIndex, slide.shapes.items)}`);
}

export async function resolveSlideShapeByIdWithXmlFallback(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  slideIndex: number,
  shapeId: string | number,
  shapeIdMap?: Record<string, string>,
): Promise<ResolvedPowerPointShapeTarget> {
  const resolved = await remapShapeIdentityAfterRoundTrip(context, slide, slideIndex, shapeId, shapeIdMap);
  return {
    shape: slide.shapes.items[resolved.shapeIndex],
    shapeId: resolved.officeShapeId,
    shapeIndex: resolved.shapeIndex,
  };
}
