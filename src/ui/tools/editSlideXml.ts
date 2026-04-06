import type { Tool } from "./types";
import { resolvePowerPointSlideIndexes } from "./powerpointContext";
import { getSlideByIndex, getSlideById } from "./powerpointNativeContent";
import { replaceSlideWithMutatedOpenXml } from "./powerpointOpenXml";
import { buildPowerPointShapeRef, parsePowerPointShapeRef, resolvePowerPointShapeRefTarget } from "./powerpointShapeRefs";
import {
  executeSlideXmlCodeInBase64Presentation,
  replaceShapeParagraphXmlInBase64Presentation,
  type SlideXmlCodeExecutionResult,
} from "./powerpointSlideXml";
import {
  getShapeTextAutoSizeSetting,
  reapplyShapeTextAutoSizeSetting,
  type PowerPointTextAutoSizeSetting,
} from "./powerpointShapeTarget";
import { roundTripRefreshHint, shouldAddRoundTripShapeTargetRefreshHint, toolFailure } from "./powerpointShared";
import { z } from "zod";

const legacyReplacementSchema = z.object({
  ref: z.string().min(1),
  paragraphsXml: z.array(z.string()),
});

const editSlideXmlArgsSchema = z.object({
  slideIndex: z.number().optional(),
  code: z.string().trim().min(1).optional(),
  autosize_shape_ids: z.array(z.union([z.string(), z.number()])).optional(),
  replacements: z.array(legacyReplacementSchema).min(1).optional(),
}).superRefine((value, context) => {
  if (!value.code && !value.replacements?.length) {
    context.addIssue({
      code: z.ZodIssueCode.custom,
      message: "Provide code or replacements.",
      path: ["code"],
    });
  }
  if (value.code && value.replacements?.length) {
    context.addIssue({
      code: z.ZodIssueCode.custom,
      message: "Provide either code or replacements, not both.",
      path: ["code"],
    });
  }
});

async function resolveSlideIndex(
  context: PowerPoint.RequestContext,
  requestedSlideIndex: number | undefined,
): Promise<number> {
  const resolved = resolvePowerPointSlideIndexes(requestedSlideIndex);
  if (typeof resolved === "number") {
    return resolved;
  }

  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  if (slides.items.length === 1) {
    return 0;
  }

  throw new Error("slideIndex is required when no active slide can be inferred from the current PowerPoint context.");
}

function normalizeAutosizeShapeIds(args: z.infer<typeof editSlideXmlArgsSchema>) {
  return (args.autosize_shape_ids || []).map((value) => String(value).trim());
}

function validateAutosizeShapeIds(autosizeShapeIds: string[]) {
  for (const xmlShapeId of autosizeShapeIds) {
    if (!/^\d+$/.test(xmlShapeId)) {
      throw new Error(`Invalid autosize_shape_ids entry ${JSON.stringify(xmlShapeId)}. Expected a numeric XML cNvPr id.`);
    }
  }
}

export const editSlideXml: Tool = {
  name: "edit_slide_xml",
  description: "General-purpose single-slide XML editor. Exports one slide as a ZIP package, exposes ppt/slides/slide1.xml for DOM-based mutation, and reimports the edited slide in one round-trip. Use for batch text edits, advanced formatting, structural shape work, and any single-slide edit that benefits from full OOXML fidelity. Prefer this over execute_office_js for text editing and formatting work.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: {
        type: "number",
        description: "Optional 0-based slide index. Required when no active slide can be inferred safely. Use this for arbitrary single-slide XML edits.",
      },
      code: {
        type: "string",
        description: "Async JavaScript function body that receives a JSZip-style single-slide package in `zip` (supporting both `zip.file(path)` and `zip.files[path]` reads), the raw slide XML string in `xml`, the parsed slide XML DOM in both `doc` and `slideXml` for ppt/slides/slide1.xml, `slidePath`, `DOMParser`, `XMLSerializer`, `escapeXml`, `namespaces`, `console`, `parseXml`, `serializeXml`, and `setResult(value)`. Returning an XML string replaces ppt/slides/slide1.xml directly.",
      },
      autosize_shape_ids: {
        type: "array",
        items: { "anyOf": [{ "type": "string" }, { "type": "number" }] },
        description: "Optional XML cNvPr shape ids whose current text auto-size settings should be preserved after the edited slide is reimported.",
      },
      replacements: {
        type: "array",
        description: "Legacy shorthand for text-only multi-shape updates on one slide. Prefer `code` for general slide XML edits.",
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
  },
  handler: async (args) => {
    const parsedArgs = editSlideXmlArgsSchema.safeParse(args ?? {});
    if (!parsedArgs.success) {
      return toolFailure(parsedArgs.error.issues[0]?.message || "Invalid arguments.");
    }

    const autosizeShapeIds = normalizeAutosizeShapeIds(parsedArgs.data);
    try {
      validateAutosizeShapeIds(autosizeShapeIds);
    } catch (error) {
      return toolFailure(error);
    }

    if (parsedArgs.data.code) {
      if (parsedArgs.data.slideIndex !== undefined && (!Number.isInteger(parsedArgs.data.slideIndex) || parsedArgs.data.slideIndex < 0)) {
        return toolFailure("slideIndex must be a non-negative integer.");
      }

      try {
        return await PowerPoint.run(async (context) => {
          const resolvedSlideIndex = await resolveSlideIndex(context, parsedArgs.data.slideIndex);
          const sourceSlide = await getSlideByIndex(context, resolvedSlideIndex);
          const rememberedAutoSizeEntries: Array<{ xmlShapeId: string; autoSizeSetting: PowerPointTextAutoSizeSetting | null }> = [];
          for (const xmlShapeId of autosizeShapeIds) {
            rememberedAutoSizeEntries.push({
              xmlShapeId,
              autoSizeSetting: await getShapeTextAutoSizeSetting(context, sourceSlide, resolvedSlideIndex, xmlShapeId),
            });
          }
          const executionHolder: { current: SlideXmlCodeExecutionResult | null } = { current: null };

          const roundTrip = await replaceSlideWithMutatedOpenXml(
            context,
            resolvedSlideIndex,
            async (base64, exportedSlide) => {
              executionHolder.current = await executeSlideXmlCodeInBase64Presentation(base64, parsedArgs.data.code!, {
                slideId: exportedSlide.id,
              });
              return executionHolder.current.mutatedBase64;
            },
          );

          const refreshedRefs = autosizeShapeIds.map((xmlShapeId) => ({
            xmlShapeId,
            ref: buildPowerPointShapeRef(roundTrip.replacementSlideId, xmlShapeId),
          }));

          if (rememberedAutoSizeEntries.length > 0) {
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
          }
          const completedExecution = executionHolder.current;

          return {
            resultType: "success",
            textResultForLlm: `Edited slide ${roundTrip.finalSlideIndex + 1} via single-slide XML round-trip.`,
            slideId: roundTrip.replacementSlideId,
            slideIndex: roundTrip.finalSlideIndex,
            slidePath: completedExecution ? completedExecution.slidePath : "ppt/slides/slide1.xml",
            result: completedExecution ? completedExecution.result : null,
            logs: completedExecution ? completedExecution.logs : [],
            hasResult: completedExecution ? completedExecution.hasResult : false,
            usedExplicitResult: completedExecution ? completedExecution.usedExplicitResult : false,
            autosize_shape_ids: autosizeShapeIds,
            refreshedRefs,
            toolTelemetry: roundTrip,
          };
        });
      } catch (error: unknown) {
        return toolFailure(error, shouldAddRoundTripShapeTargetRefreshHint(error) ? roundTripRefreshHint() : undefined);
      }
    }

    const replacements = parsedArgs.data.replacements;
    if (!replacements?.length) {
      return toolFailure("Provide code or replacements.");
    }

    let parsedRefs: Array<ReturnType<typeof parsePowerPointShapeRef>>;
    try {
      parsedRefs = replacements.map((replacement) => parsePowerPointShapeRef(replacement.ref));
    } catch (error) {
      return toolFailure(error);
    }

    const slideIds = new Set(parsedRefs.map((replacement) => replacement.slideId));
    if (slideIds.size !== 1) {
      return toolFailure("All replacements must target the same slide.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        const firstTarget = await resolvePowerPointShapeRefTarget(context, replacements[0].ref);
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
            replacements.map((replacement, index) => ({
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

          const updatedRefs = parsedRefs.map((replacement) => ({
            ref: buildPowerPointShapeRef(roundTrip.replacementSlideId, replacement.xmlShapeId),
            slideId: roundTrip.replacementSlideId,
            xmlShapeId: replacement.xmlShapeId,
          }));

          return {
            resultType: "success",
            textResultForLlm: `Updated ${updatedRefs.length} shapes on slide ${roundTrip.finalSlideIndex + 1}.`,
            slideId: roundTrip.replacementSlideId,
            slideIndex: roundTrip.finalSlideIndex,
            replacements: updatedRefs,
            toolTelemetry: roundTrip,
          };
      });
    } catch (error: unknown) {
      return toolFailure(error, shouldAddRoundTripShapeTargetRefreshHint(error) ? roundTripRefreshHint() : undefined);
    }
  },
};
