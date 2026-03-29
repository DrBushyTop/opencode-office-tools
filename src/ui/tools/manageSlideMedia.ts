import type { Tool } from "./types";
import { resolvePowerPointTargetingArgs } from "./powerpointContext";
import { createImageRectangle, fetchImageUrlAsBase64, getShapeBounds, getSlideByIndex } from "./powerpointNativeContent";
import { resolveSlideShapeByIdWithXmlFallback } from "./powerpointShapeTarget";
import { toolFailure } from "./powerpointShared";
import { z } from "zod";

type ManageSlideMediaAction = "insertImage" | "replaceImage" | "deleteImage";

const manageSlideMediaArgsSchema = z.object({
  action: z.enum(["insertImage", "replaceImage", "deleteImage"]),
  slideIndex: z.number().optional(),
  shapeId: z.union([z.string(), z.number()]).optional(),
  imageUrl: z.string().optional(),
  left: z.number().optional(),
  top: z.number().optional(),
  width: z.number().optional(),
  height: z.number().optional(),
  name: z.string().optional(),
});

type ManageSlideMediaArgs = z.infer<typeof manageSlideMediaArgsSchema>;

export const manageSlideMedia: Tool = {
  name: "manage_slide_media",
  description: "Insert, replace, or delete editable PowerPoint image shapes on a slide.",
  parameters: {
    type: "object",
    properties: {
      action: { type: "string", enum: ["insertImage", "replaceImage", "deleteImage"], description: "Media action to perform." },
      slideIndex: { type: "number", description: "0-based slide index. Defaults to the active slide when available." },
      shapeId: { anyOf: [{ type: "string" }, { type: "number" }], description: "Existing image shape id for replaceImage or deleteImage." },
      imageUrl: { type: "string", description: "Source image URL for insertImage or replaceImage." },
      left: { type: "number" },
      top: { type: "number" },
      width: { type: "number" },
      height: { type: "number" },
      name: { type: "string", description: "Optional shape name for the inserted image container." },
    },
    required: ["action"],
  },
  handler: async (args) => {
    const parsedArgs = manageSlideMediaArgsSchema.safeParse(args);
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }
    const media = resolvePowerPointTargetingArgs(parsedArgs.data as ManageSlideMediaArgs);
    if (!Number.isInteger(media.slideIndex) || (media.slideIndex as number) < 0) {
      return toolFailure("slideIndex must be a non-negative integer.");
    }
    if ((media.action === "insertImage" || media.action === "replaceImage") && !media.imageUrl) {
      return toolFailure("imageUrl is required for insertImage and replaceImage.");
    }
    if ((media.action === "replaceImage" || media.action === "deleteImage") && media.shapeId === undefined) {
      return toolFailure("shapeId is required for replaceImage and deleteImage.");
    }

    const slideIndex = media.slideIndex as number;

    try {
      return await PowerPoint.run(async (context) => {
        const slide = await getSlideByIndex(context, slideIndex);

        if (media.action === "insertImage") {
          const imageBase64 = await fetchImageUrlAsBase64(media.imageUrl!);
          const created = createImageRectangle(slide, {
            left: media.left ?? 60,
            top: media.top ?? 80,
            width: media.width ?? 280,
            height: media.height ?? 180,
            name: media.name || "Image",
            imageBase64,
          });
          created.load(["id", "name"]);
          await context.sync();
          return {
            resultType: "success",
            textResultForLlm: `Inserted image ${created.id} on slide ${slideIndex + 1}.`,
            slideIndex,
            shapeId: created.id,
            toolTelemetry: { slideIndex, shapeId: created.id },
          };
        }

        const resolved = await resolveSlideShapeByIdWithXmlFallback(context, slide, slideIndex, media.shapeId!);

        if (media.action === "deleteImage") {
          resolved.shape.delete();
          await context.sync();
          return {
            resultType: "success",
            textResultForLlm: `Deleted image shape ${resolved.shapeId} from slide ${slideIndex + 1}.`,
            slideIndex,
            shapeId: resolved.shapeId,
            toolTelemetry: { slideIndex, shapeId: resolved.shapeId },
          };
        }

        const bounds = await getShapeBounds(resolved.shape, context);
        const imageBase64 = await fetchImageUrlAsBase64(media.imageUrl!);
        resolved.shape.delete();
        const created = createImageRectangle(slide, {
          left: media.left ?? bounds.left,
          top: media.top ?? bounds.top,
          width: media.width ?? bounds.width,
          height: media.height ?? bounds.height,
          name: media.name || bounds.name || "Image",
          imageBase64,
        });
        created.load(["id", "name"]);
        await context.sync();
        return {
          resultType: "success",
          textResultForLlm: `Replaced image shape ${resolved.shapeId} on slide ${slideIndex + 1}.`,
          slideIndex,
          shapeId: created.id,
          replacedShapeId: resolved.shapeId,
          toolTelemetry: { slideIndex, replacedShapeId: resolved.shapeId, shapeId: created.id },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
