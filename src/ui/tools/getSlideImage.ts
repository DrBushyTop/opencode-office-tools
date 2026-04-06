import type { Tool } from "./types";
import { z } from "zod";
import { describeError } from "./powerpointShared";

const getSlideImageArgsSchema = z.object({
  slideIndex: z.number(),
  width: z.number().optional(),
});

function buildCaptureFailure(message: string, error = message) {
  return { textResultForLlm: message, resultType: "failure" as const, error, toolTelemetry: {} };
}

function formatOutOfRangeMessage(slideIndex: number, slideCount: number) {
  return `Invalid slideIndex ${slideIndex}. Must be 0-${slideCount - 1} (current slide count: ${slideCount})`;
}

function validateCaptureRequest(args: unknown) {
  const parsedArgs = getSlideImageArgsSchema.safeParse(args);
  if (!parsedArgs.success) {
    return { ok: false as const, failure: buildCaptureFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.") };
  }

  const { slideIndex, width = 800 } = parsedArgs.data;
  if (!Number.isInteger(slideIndex) || slideIndex < 0) {
    return { ok: false as const, failure: buildCaptureFailure("slideIndex must be a non-negative integer.") };
  }
  if (!Number.isFinite(width) || width <= 0) {
    return { ok: false as const, failure: buildCaptureFailure("width must be a positive number.") };
  }

  return { ok: true as const, data: { slideIndex, width } };
}

async function exportSlideImage(context: PowerPoint.RequestContext, slideIndex: number, width: number) {
  const deck = context.presentation.slides;
  deck.load("items");
  await context.sync();

  if (slideIndex >= deck.items.length) {
    return buildCaptureFailure(formatOutOfRangeMessage(slideIndex, deck.items.length), "Invalid slideIndex");
  }

  const request = deck.items[slideIndex].getImageAsBase64({ width });
  await context.sync();

  return {
    textResultForLlm: `Rendered slide ${slideIndex + 1} of ${deck.items.length} as a ${width}px PNG snapshot.`,
    binaryResultsForLlm: [
      {
        data: request.value,
        mimeType: "image/png",
        type: "image",
        description: `Slide ${slideIndex + 1} of ${deck.items.length}`,
      },
    ],
    resultType: "success" as const,
    toolTelemetry: {},
  };
}

function isImageExportUnavailable(error: unknown) {
  const message = describeError(error);
  return message.includes("getImageAsBase64") || (error as { code?: string } | null)?.code === "InvalidOperation";
}

export const getSlideImage: Tool = {
  name: "get_slide_image",
  description: "Render one PowerPoint slide as a PNG snapshot. Useful when you need to inspect layout, spacing, typography, or colors before editing nearby content.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "0-based slide index. Use 0 for first slide, 1 for second, etc.",
      },
      width: {
        type: "number",
        description: "Optional width in pixels for the image. Aspect ratio is preserved. Default is 800.",
      },
    },
    required: ["slideIndex"],
  },
  handler: async (args) => {
    const request = validateCaptureRequest(args);
    if (!request.ok) return request.failure;

    try {
      return await PowerPoint.run((context) => exportSlideImage(context, request.data.slideIndex, request.data.width));
    } catch (error: unknown) {
      if (isImageExportUnavailable(error)) {
      return buildCaptureFailure(
          "This PowerPoint host cannot export slide images. Use a recent PowerPoint build on Windows, Mac, or the web and try again.",
          "API not available",
        );
      }
      return buildCaptureFailure(describeError(error));
    }
  },
};
