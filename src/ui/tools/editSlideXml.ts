import type { Tool } from "./types";
import { getSlideById } from "./powerpointNativeContent";
import { replaceSlideWithMutatedOpenXml } from "./powerpointOpenXml";
import { buildPowerPointShapeRef, parsePowerPointShapeRef, resolvePowerPointShapeRefTarget } from "./powerpointShapeRefs";
import { replaceShapeParagraphXmlInBase64Presentation } from "./powerpointSlideXml";
import { getShapeTextAutoSizeSetting, reapplyShapeTextAutoSizeSetting } from "./powerpointShapeTarget";
import { roundTripRefreshHint, shouldAddRoundTripShapeTargetRefreshHint, toolFailure } from "./powerpointShared";
import { z } from "zod";

const editSlideXmlArgsSchema = z.object({
  replacements: z.array(z.object({
    ref: z.string().min(1),
    paragraphsXml: z.array(z.string()),
  })).min(1),
});

export const editSlideXml: Tool = {
  name: "edit_slide_xml",
  description: "Batch-edit multiple PowerPoint text shapes on one slide in a single Open XML round-trip using stable shape refs.",
  parameters: {
    type: "object",
    properties: {
      replacements: {
        type: "array",
        items: {
          type: "object",
          properties: {
            ref: { type: "string" },
            paragraphsXml: { type: "array", items: { type: "string" } },
          },
          required: ["ref", "paragraphsXml"],
        },
      },
    },
    required: ["replacements"],
  },
  handler: async (args) => {
    const parsedArgs = editSlideXmlArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    let parsedRefs: Array<ReturnType<typeof parsePowerPointShapeRef>>;
    try {
      parsedRefs = parsedArgs.data.replacements.map((replacement) => parsePowerPointShapeRef(replacement.ref));
    } catch (error) {
      return toolFailure(error);
    }

    const slideIds = new Set(parsedRefs.map((replacement) => replacement.slideId));
    if (slideIds.size !== 1) {
      return toolFailure("All replacements must target the same slide.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const firstTarget = await resolvePowerPointShapeRefTarget(context, parsedArgs.data.replacements[0].ref);
        const rememberedAutoSizeEntries: Array<{ xmlShapeId: string; autoSizeSetting: Awaited<ReturnType<typeof getShapeTextAutoSizeSetting>> }> = [];
        for (const replacement of parsedRefs) {
          rememberedAutoSizeEntries.push({
            xmlShapeId: replacement.xmlShapeId,
            autoSizeSetting: await getShapeTextAutoSizeSetting(context, firstTarget.slide, firstTarget.slideIndex, replacement.xmlShapeId),
          });
        }

        const roundTrip = await replaceSlideWithMutatedOpenXml(
          context,
          firstTarget.slideIndex,
          (base64) => replaceShapeParagraphXmlInBase64Presentation(
            base64,
            parsedArgs.data.replacements.map((replacement, index) => ({
              target: {
                slideId: parsedRefs[index].slideId,
                xmlShapeId: parsedRefs[index].xmlShapeId,
                ref: parsedRefs[index].ref,
              },
              paragraphsXml: replacement.paragraphsXml,
            })),
            { slideId: firstTarget.slideId },
          ),
        );

        try {
          const { slide: replacementSlide } = await getSlideById(context, roundTrip.replacementSlideId);
          for (const entry of rememberedAutoSizeEntries) {
            await reapplyShapeTextAutoSizeSetting(
              context,
              replacementSlide,
              roundTrip.finalSlideIndex,
              entry.xmlShapeId,
              entry.autoSizeSetting,
            );
          }
        } catch {
          // Best-effort only.
        }

        const replacements = parsedRefs.map((replacement) => ({
          ref: buildPowerPointShapeRef(roundTrip.replacementSlideId, replacement.xmlShapeId),
          slideId: roundTrip.replacementSlideId,
          xmlShapeId: replacement.xmlShapeId,
        }));

        return {
          resultType: "success",
          textResultForLlm: `Updated ${replacements.length} shapes on slide ${roundTrip.finalSlideIndex + 1}.`,
          slideId: roundTrip.replacementSlideId,
          slideIndex: roundTrip.finalSlideIndex,
          replacements,
          toolTelemetry: roundTrip,
        };
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripShapeTargetRefreshHint(error) ? roundTripRefreshHint() : undefined);
    }
  },
};
