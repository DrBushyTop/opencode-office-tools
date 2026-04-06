import type { Tool } from "./types";
import pptxgen from "pptxgenjs";
import { isPowerPointRequirementSetSupported } from "./powerpointShared";
import { z } from "zod";

const addSlideFromCodeArgsSchema = z.object({
  code: z.string(),
  replaceSlideIndex: z.number().optional(),
});

type AddSlideFromCodeArgs = z.infer<typeof addSlideFromCodeArgsSchema>;

type SlidePackage = {
  base64: string;
  widthInches: number;
  heightInches: number;
};

type InsertPlan = {
  mode: "append" | "replace";
  targetSlideId?: string;
  removeAfterInsertIndex?: number;
};

function failureResult(message: string, error = message, includeStack?: string) {
  return {
    textResultForLlm: includeStack ? `${message}\n\nStack: ${includeStack}` : message,
    resultType: "failure" as const,
    error,
    toolTelemetry: {},
  };
}

function successResult(args: AddSlideFromCodeArgs) {
  return args.replaceSlideIndex !== undefined
    ? `Replaced slide ${args.replaceSlideIndex + 1} with generated content.`
    : "Inserted a generated slide into the presentation.";
}

export function normalizeAddSlideCode(input: string) {
  return String(input || "")
    .replace(/^\s*import\s+.+?from\s+["']pptxgenjs["'];?\s*$/gm, "")
    .replace(/^\s*const\s+\w+\s*=\s*require\(["']pptxgenjs["']\);?\s*$/gm, "")
    .replace(/^\s*(?:const|let|var)\s+pptx\s*=\s*new\s+\w+\s*\(\s*\)\s*;?\s*$/gm, "")
    .replace(/\b(?:const|let|var)\s+slide\s*=\s*pptx\.addSlide\(\s*\)\s*;?/g, "")
    .replace(/\b(?:const|let|var)\s+(\w+)\s*=\s*pptx\.addSlide\(\s*\)\s*;?/g, "const $1 = slide;")
    .trim();
}

export function buildGeneratedSlide(code: string, slide: any, pptx: any) {
  const runtime = {
    slide,
    pptx,
    pptxgen,
    ShapeType: pptx.ShapeType,
    AlignH: pptx.AlignH,
    AlignV: pptx.AlignV,
  };

  const compiled = new Function("runtime", `with (runtime) { ${normalizeAddSlideCode(code)} }`);
  compiled(runtime);
}

async function readDeckDimensions() {
  let widthInches = 13.333;
  let heightInches = 7.5;

  if (!isPowerPointRequirementSetSupported("1.10")) {
    return { widthInches, heightInches };
  }

  try {
    await PowerPoint.run(async (context) => {
      const pageSetup = context.presentation.pageSetup;
      pageSetup.load(["slideWidth", "slideHeight"]);
      await context.sync();
      widthInches = pageSetup.slideWidth / 72;
      heightInches = pageSetup.slideHeight / 72;
    });
  } catch {
    // Keep the fallback widescreen size when the host cannot report page setup.
  }

  return { widthInches, heightInches };
}

function validateAddSlideArgs(args: unknown) {
  const parsed = addSlideFromCodeArgsSchema.safeParse(args);
  if (!parsed.success) {
    return { ok: false as const, failure: failureResult(parsed.error.issues[0]?.message || "Invalid arguments.") };
  }

  if (parsed.data.replaceSlideIndex !== undefined && (!Number.isInteger(parsed.data.replaceSlideIndex) || parsed.data.replaceSlideIndex < 0)) {
    return { ok: false as const, failure: failureResult("replaceSlideIndex must be a non-negative integer.") };
  }

  return { ok: true as const, data: parsed.data };
}

async function buildSlidePackage(args: AddSlideFromCodeArgs): Promise<SlidePackage> {
  const dimensions = await readDeckDimensions();
  const deck = new pptxgen();
  deck.defineLayout({ name: "OPENCODE_RUNTIME_LAYOUT", width: dimensions.widthInches, height: dimensions.heightInches });
  deck.layout = "OPENCODE_RUNTIME_LAYOUT";
  const slide = deck.addSlide();

  buildGeneratedSlide(args.code, slide, deck);
  const base64 = await deck.write({ outputType: "base64" }) as string;
  return { base64, ...dimensions };
}

async function loadSlides(context: PowerPoint.RequestContext) {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  return slides;
}

async function buildInsertPlan(context: PowerPoint.RequestContext, replaceSlideIndex?: number): Promise<InsertPlan> {
  const slides = await loadSlides(context);

  if (replaceSlideIndex === undefined) {
    if (slides.items.length === 0) {
      return { mode: "append" };
    }

    const previousSlide = slides.items[slides.items.length - 1];
    previousSlide.load("id");
    await context.sync();
    return { mode: "append", targetSlideId: previousSlide.id };
  }

  if (replaceSlideIndex >= slides.items.length) {
    throw new Error(`Invalid replaceSlideIndex ${replaceSlideIndex}. Must be 0-${slides.items.length - 1} (current slide count: ${slides.items.length})`);
  }

  if (replaceSlideIndex === 0) {
    return { mode: "replace", removeAfterInsertIndex: 1 };
  }

  const anchor = slides.items[replaceSlideIndex - 1];
  anchor.load("id");
  await context.sync();
  return { mode: "replace", targetSlideId: anchor.id, removeAfterInsertIndex: replaceSlideIndex + 1 };
}

async function applyInsertPlan(context: PowerPoint.RequestContext, slidePackage: SlidePackage, plan: InsertPlan) {
  context.presentation.insertSlidesFromBase64(slidePackage.base64, {
    formatting: PowerPoint.InsertSlideFormatting.useDestinationTheme,
    ...(plan.targetSlideId ? { targetSlideId: plan.targetSlideId } : {}),
  });
  await context.sync();

  if (plan.mode !== "replace" || plan.removeAfterInsertIndex === undefined) {
    return;
  }

  const slides = await loadSlides(context);
  if (plan.removeAfterInsertIndex < slides.items.length) {
    slides.items[plan.removeAfterInsertIndex].delete();
    await context.sync();
  }
}

export const addSlideFromCode: Tool = {
  name: "add_slide_from_code",
  description: `Build a PowerPoint slide from a PptxGenJS recipe, then import that generated slide into the live deck.

Use this when native slide-editing tools are not enough and you need a programmatic fallback.

Accepted code styles:
- direct calls on the provided \`slide\` object
- fuller snippets that refer to \`pptx\`, \`pptxgen\`, or \`pptx.addSlide()\`

Runtime objects available to your code:
- \`slide\`
- \`pptx\`
- \`pptxgen\`
- \`ShapeType\`
- \`AlignH\`
- \`AlignV\`

The generated slide is sized to the current deck when the host exposes page setup information.`,
  parameters: {
    type: "object",
    properties: {
      code: {
        type: "string",
        description: "JavaScript statements that assemble one slide with PptxGenJS.",
      },
      replaceSlideIndex: {
        type: "number",
        description: "Optional 0-based slide index to replace after the generated slide is imported.",
      },
    },
    required: ["code"],
  },
  handler: async (args) => {
    const validated = validateAddSlideArgs(args);
    if (!validated.ok) return validated.failure;

    try {
      const slidePackage = await buildSlidePackage(validated.data);

      try {
        await PowerPoint.run(async (context) => {
          const plan = await buildInsertPlan(context, validated.data.replaceSlideIndex);
          await applyInsertPlan(context, slidePackage, plan);
        });
      } catch (insertError: any) {
        return failureResult(`Failed to insert slide: ${insertError.message}`, insertError.message);
      }

      return successResult(validated.data);
    } catch (error: any) {
      const text = String(error?.message || error || "Unknown error");
      if (/write|generate/i.test(text)) {
        return failureResult(`Presentation generation failed: ${text}`, text);
      }
      if (/syntax|reference|type|is not defined|unexpected token/i.test(text)) {
        return failureResult(`The supplied slide recipe failed while running: ${text}`, text, error?.stack);
      }
      return failureResult(`Unexpected add_slide_from_code failure: ${text}`, text, error?.stack);
    }
  },
};
