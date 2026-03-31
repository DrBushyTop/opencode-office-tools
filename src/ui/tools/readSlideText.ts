import type { Tool } from "./types";
import { getShapeParagraphXmlByTarget } from "./powerpointSlideXml";
import { resolvePowerPointShapeRefTarget } from "./powerpointShapeRefs";
import { roundTripRefreshHint, shouldAddRoundTripShapeTargetRefreshHint, toolFailure } from "./powerpointShared";
import { z } from "zod";

const readSlideTextArgsSchema = z.object({
  ref: z.string().min(1),
});

export const readSlideText: Tool = {
  name: "read_slide_text",
  description: "Read the raw <a:p> paragraph XML from one PowerPoint shape addressed by a stable shape ref.",
  parameters: {
    type: "object",
    properties: {
      ref: {
        type: "string",
        description: "Stable shape ref in the format slide-id:<slideId>/shape:<xmlShapeId>.",
      },
    },
    required: ["ref"],
  },
  handler: async (args) => {
    const parsedArgs = readSlideTextArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const target = await resolvePowerPointShapeRefTarget(context, parsedArgs.data.ref);
        const paragraphsXml = getShapeParagraphXmlByTarget(target);

        return {
          resultType: "success",
          textResultForLlm: `Read ${paragraphsXml.length} paragraphs from shape ${target.ref}.`,
          ref: target.ref,
          slideId: target.slideId,
          xmlShapeId: target.xmlShapeId,
          slideIndex: target.slideIndex,
          paragraphsXml,
          toolTelemetry: {},
        };
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripShapeTargetRefreshHint(error) ? roundTripRefreshHint() : undefined);
    }
  },
};
