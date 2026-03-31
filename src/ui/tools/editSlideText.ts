import type { Tool } from "./types";
import { getSlideById } from "./powerpointNativeContent";
import { replaceSlideWithMutatedOpenXml } from "./powerpointOpenXml";
import { buildPowerPointShapeRef, resolvePowerPointShapeRefTarget } from "./powerpointShapeRefs";
import { replaceShapeParagraphXmlInBase64Presentation } from "./powerpointSlideXml";
import { getShapeTextAutoSizeSetting, reapplyShapeTextAutoSizeSetting } from "./powerpointShapeTarget";
import { roundTripRefreshHint, shouldAddRoundTripShapeTargetRefreshHint, toolFailure } from "./powerpointShared";
import { z } from "zod";

const editSlideTextArgsSchema = z.object({
  ref: z.string().min(1),
  paragraphsXml: z.array(z.string()),
});

export const editSlideText: Tool = {
  name: "edit_slide_text",
  description: "Preferred text-editing tool for one existing PowerPoint text shape. Replaces raw paragraph XML in place via one slide Open XML round-trip while preserving text body properties.",
  parameters: {
    type: "object",
    properties: {
      ref: {
        type: "string",
        description: "Stable shape ref in the format slide-id:<slideId>/shape:<xmlShapeId>.",
      },
      paragraphsXml: {
        type: "array",
        items: { type: "string" },
        description: "Replacement raw <a:p> paragraph XML for the targeted shape.",
      },
    },
    required: ["ref", "paragraphsXml"],
  },
  handler: async (args) => {
    const parsedArgs = editSlideTextArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const target = await resolvePowerPointShapeRefTarget(context, parsedArgs.data.ref);
        const rememberedAutoSize = await getShapeTextAutoSizeSetting(context, target.slide, target.slideIndex, target.xmlShapeId);
        const roundTrip = await replaceSlideWithMutatedOpenXml(
          context,
          target.slideIndex,
          (base64) => replaceShapeParagraphXmlInBase64Presentation(base64, [{
            target: {
              slideId: target.slideId,
              xmlShapeId: target.xmlShapeId,
              ref: target.ref,
            },
            paragraphsXml: parsedArgs.data.paragraphsXml,
          }], { slideId: target.slideId }),
        );

        let autoSizeRetained = false;
        if (rememberedAutoSize === "AutoSizeShapeToFitText") {
          try {
            const { slide: replacementSlide } = await getSlideById(context, roundTrip.replacementSlideId);
            autoSizeRetained = await reapplyShapeTextAutoSizeSetting(
              context,
              replacementSlide,
              roundTrip.finalSlideIndex,
              target.xmlShapeId,
              rememberedAutoSize,
            );
          } catch {
            autoSizeRetained = false;
          }
        }

        const refreshedRef = buildPowerPointShapeRef(roundTrip.replacementSlideId, target.xmlShapeId);
        return {
          resultType: "success",
          textResultForLlm: `Updated shape text on slide ${roundTrip.finalSlideIndex + 1}.`,
          ref: refreshedRef,
          slideId: roundTrip.replacementSlideId,
          xmlShapeId: target.xmlShapeId,
          slideIndex: roundTrip.finalSlideIndex,
          autoSizeRetained,
          toolTelemetry: roundTrip,
        };
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripShapeTargetRefreshHint(error) ? roundTripRefreshHint() : undefined);
    }
  },
};
