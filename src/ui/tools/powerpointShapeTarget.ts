import { findSlideShapeIndexByXmlShapeIdInBase64Presentation } from "./powerpointOpenXml";
import { formatAvailableShapeTargets, isPowerPointRequirementSetSupported } from "./powerpointShared";
import type { PowerPointShapeIdentifier } from "./powerpointContext";

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

export type PowerPointTextAutoSizeSetting = "AutoSizeNone" | "AutoSizeTextToFitShape" | "AutoSizeShapeToFitText";

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
  requestedShapeId: PowerPointShapeIdentifier,
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

  if (exportedBase64) {
    const xmlMatchIndex = findSlideShapeIndexByXmlShapeIdInBase64Presentation(exportedBase64, requestedId);
    if (xmlMatchIndex >= 0) {
      const xmlMatch = slide.shapes.items[xmlMatchIndex];
      if (xmlMatch) {
        return {
          officeShapeId: xmlMatch.id,
          xmlShapeId: requestedId,
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
  shapeId: PowerPointShapeIdentifier,
): Promise<ResolvedPowerPointShapeTarget> {
  const resolved = await remapShapeIdentityAfterRoundTrip(context, slide, slideIndex, shapeId);
  return {
    shape: slide.shapes.items[resolved.shapeIndex],
    shapeId: resolved.officeShapeId,
    shapeIndex: resolved.shapeIndex,
  };
}

export async function getShapeTextAutoSizeSetting(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  slideIndex: number,
  shapeId: PowerPointShapeIdentifier,
): Promise<PowerPointTextAutoSizeSetting | null> {
  const resolved = await resolveSlideShapeByIdWithXmlFallback(context, slide, slideIndex, shapeId);
  const frame = resolved.shape.getTextFrameOrNullObject();
  frame.load(["isNullObject", "autoSizeSetting"]);
  await context.sync();

  return frame.isNullObject ? null : (frame.autoSizeSetting as PowerPointTextAutoSizeSetting | null);
}

export async function reapplyShapeTextAutoSizeSetting(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  slideIndex: number,
  shapeId: PowerPointShapeIdentifier,
  autoSizeSetting: PowerPointTextAutoSizeSetting | null | undefined,
): Promise<boolean> {
  if (autoSizeSetting !== "AutoSizeShapeToFitText") {
    return false;
  }

  const resolved = await resolveSlideShapeByIdWithXmlFallback(context, slide, slideIndex, shapeId);
  const frame = resolved.shape.getTextFrameOrNullObject();
  frame.load("isNullObject");
  await context.sync();
  if (frame.isNullObject) {
    return false;
  }

  frame.autoSizeSetting = autoSizeSetting;
  await context.sync();
  return true;
}
