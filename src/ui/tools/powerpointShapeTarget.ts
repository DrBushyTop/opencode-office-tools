import { findSlideShapeIndexByXmlShapeIdInBase64Presentation } from "./powerpointOpenXml";
import { formatAvailableShapeTargets, isPowerPointRequirementSetSupported } from "./powerpointShared";

export interface ResolvedPowerPointShapeTarget {
  shape: PowerPoint.Shape;
  shapeId: string;
  shapeIndex: number;
}

export async function resolveSlideShapeByIdWithXmlFallback(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  slideIndex: number,
  shapeId: string | number,
): Promise<ResolvedPowerPointShapeTarget> {
  const requestedShapeId = String(shapeId);

  slide.shapes.load("items/id,name");
  await context.sync();

  const officeMatchIndex = slide.shapes.items.findIndex((shape) => shape.id === requestedShapeId);
  if (officeMatchIndex >= 0) {
    return {
      shape: slide.shapes.items[officeMatchIndex],
      shapeId: requestedShapeId,
      shapeIndex: officeMatchIndex,
    };
  }

  if (isPowerPointRequirementSetSupported("1.8")) {
    const exported = slide.exportAsBase64();
    await context.sync();

    const xmlMatchIndex = findSlideShapeIndexByXmlShapeIdInBase64Presentation(exported.value, requestedShapeId);
    if (xmlMatchIndex >= 0) {
      const xmlMatch = slide.shapes.items[xmlMatchIndex];
      if (xmlMatch) {
        return {
          shape: xmlMatch,
          shapeId: xmlMatch.id,
          shapeIndex: xmlMatchIndex,
        };
      }
    }
  }

  throw new Error(`Shape ${requestedShapeId} was not found on slide ${slideIndex + 1}. ${formatAvailableShapeTargets(slideIndex, slide.shapes.items)}`);
}
