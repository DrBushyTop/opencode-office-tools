import type { Tool } from "./types";
import { resolvePowerPointTargetingArgs } from "./powerpointContext";
import { applyTemplateBindingsToSlide, createSlideWithLayout } from "./createSlideFromTemplate";
import { getSlideByIndex } from "./powerpointNativeContent";
import { loadShapeSummaries, loadThemeColors, pickThemeColor, toolFailure } from "./powerpointShared";

type LayoutType = "timeline" | "phase_plan" | "process_flow" | "comparison_grid" | "estimate_summary";

interface LayoutItem {
  title: string;
  subtitle?: string;
  value?: string;
  colorToken?: string;
}

interface InsertBusinessLayoutArgs {
  slideIndex?: number;
  layoutType: LayoutType;
  title?: string;
  items: LayoutItem[];
  themeMode?: "deck" | "custom";
}

function layoutKeywords(layoutType: LayoutType) {
  if (layoutType === "timeline") return ["timeline", "process", "roadmap"];
  if (layoutType === "phase_plan") return ["phase", "plan", "agenda"];
  if (layoutType === "comparison_grid") return ["comparison", "two objects", "table"];
  if (layoutType === "estimate_summary") return ["table", "summary", "content"];
  return ["process", "flow", "content"];
}

async function findTemplateLayout(context: PowerPoint.RequestContext, layoutType: LayoutType) {
  const keywords = layoutKeywords(layoutType);
  const masters = context.presentation.slideMasters;
  masters.load("items");
  await context.sync();
  for (const master of masters.items) {
    master.load(["id", "layouts/items/id", "layouts/items/name"]);
  }
  await context.sync();

  for (const master of masters.items) {
    for (const layout of master.layouts.items) {
      const name = String(layout.name || "").toLowerCase();
      if (keywords.some((keyword) => name.includes(keyword))) {
        return { slideMasterId: master.id, layoutId: layout.id };
      }
    }
  }
  return null;
}

async function drawBusinessLayout(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  layoutType: LayoutType,
  items: LayoutItem[],
  title: string | undefined,
) {
  const masters = context.presentation.slideMasters;
  masters.load("items");
  await context.sync();
  const colors = masters.items[0] ? await loadThemeColors(context, masters.items[0]) : undefined;

  slide.shapes.load("items");
  await context.sync();
  const shapeSummaries = await loadShapeSummaries(context, slide.shapes.items, { includeText: true, includeFormatting: false, includeTableValues: false });
  const titlePlaceholder = shapeSummaries.find((shape) => shape.placeholderType === "Title" || shape.placeholderType === "CenterTitle");
  const bodyPlaceholder = shapeSummaries.find((shape) => shape.placeholderType === "Body" || shape.placeholderType === "Content");
  const frame = bodyPlaceholder
    ? { left: bodyPlaceholder.left, top: bodyPlaceholder.top, width: bodyPlaceholder.width, height: bodyPlaceholder.height }
    : { left: 60, top: titlePlaceholder ? titlePlaceholder.top + titlePlaceholder.height + 24 : 110, width: 580, height: 280 };

  if (title && titlePlaceholder) {
    const target = slide.shapes.items[titlePlaceholder.index].getTextFrameOrNullObject();
    target.load("isNullObject");
    await context.sync();
    if (!target.isNullObject) {
      target.textRange.text = title;
    }
  } else if (title) {
    const heading = slide.shapes.addTextBox(title, { left: frame.left, top: frame.top - 48, width: frame.width, height: 30 });
    heading.textFrame.textRange.font.size = 24;
    heading.textFrame.textRange.font.bold = true;
    heading.lineFormat.visible = false;
  }

  const accent = pickThemeColor(colors, "Accent1", "#1d3557");
  const accentTwo = pickThemeColor(colors, "Accent2", "#457b9d");
  const accentThree = pickThemeColor(colors, "Accent3", "#2a9d8f");
  const palette = [accent, accentTwo, accentThree, pickThemeColor(colors, "Accent4", "#e9c46a"), pickThemeColor(colors, "Accent5", "#f4a261")];

  if (layoutType === "timeline" || layoutType === "process_flow") {
    const gap = 14;
    const cardWidth = (frame.width - gap * Math.max(items.length - 1, 0)) / Math.max(items.length, 1);
    const line = slide.shapes.addLine(PowerPoint.ConnectorType.straight, { left: frame.left, top: frame.top + 82, width: frame.width, height: 0 });
    line.lineFormat.color = accent;
    line.lineFormat.weight = 3;
    for (const [index, item] of items.entries()) {
      const left = frame.left + index * (cardWidth + gap);
      const circle = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, { left: left + cardWidth / 2 - 14, top: frame.top + 68, width: 28, height: 28 });
      circle.fill.setSolidColor(palette[index % palette.length]);
      circle.lineFormat.visible = false;
      const card = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, { left, top: frame.top + 112, width: cardWidth, height: 92 });
      card.fill.setSolidColor("#f7f8fb");
      card.lineFormat.color = palette[index % palette.length];
      card.lineFormat.weight = 1.5;
      const text = slide.shapes.addTextBox(`${item.title}${item.subtitle ? `\n${item.subtitle}` : ""}${item.value ? `\n${item.value}` : ""}`, { left: left + 10, top: frame.top + 124, width: cardWidth - 20, height: 72 });
      text.textFrame.textRange.font.size = 12;
      text.lineFormat.visible = false;
    }
    await context.sync();
    return;
  }

  if (layoutType === "comparison_grid") {
    const columns = Math.min(items.length, 3);
    const rows = Math.ceil(items.length / columns);
    const gap = 14;
    const cardWidth = (frame.width - gap * (columns - 1)) / columns;
    const cardHeight = (frame.height - gap * (rows - 1)) / rows;
    items.forEach((item, index) => {
      const column = index % columns;
      const row = Math.floor(index / columns);
      const card = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, {
        left: frame.left + column * (cardWidth + gap),
        top: frame.top + row * (cardHeight + gap),
        width: cardWidth,
        height: cardHeight,
      });
      card.fill.setSolidColor("#f7f8fb");
      card.lineFormat.color = palette[index % palette.length];
      const text = slide.shapes.addTextBox(`${item.title}${item.subtitle ? `\n${item.subtitle}` : ""}${item.value ? `\n\n${item.value}` : ""}`, {
        left: frame.left + column * (cardWidth + gap) + 12,
        top: frame.top + row * (cardHeight + gap) + 12,
        width: cardWidth - 24,
        height: cardHeight - 24,
      });
      text.textFrame.textRange.font.size = 13;
      text.lineFormat.visible = false;
    });
    await context.sync();
    return;
  }

  if (layoutType === "estimate_summary") {
    slide.shapes.addTable(items.length + 1, 3, {
      left: frame.left,
      top: frame.top,
      width: frame.width,
      height: Math.min(frame.height, 40 * (items.length + 1)),
      values: [["Item", "Detail", "Value"], ...items.map((item) => [item.title, item.subtitle || "", item.value || ""])],
    });
    await context.sync();
    return;
  }

  const rowHeight = frame.height / Math.max(items.length, 1);
  items.forEach((item, index) => {
    const y = frame.top + index * rowHeight;
    const chip = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundRectangle, { left: frame.left, top: y + 8, width: 120, height: rowHeight - 16 });
    chip.fill.setSolidColor(palette[index % palette.length]);
    chip.lineFormat.visible = false;
    const text = slide.shapes.addTextBox(`${item.title}${item.subtitle ? `\n${item.subtitle}` : ""}${item.value ? `\n${item.value}` : ""}`, { left: frame.left + 138, top: y + 8, width: frame.width - 138, height: rowHeight - 16 });
    text.textFrame.textRange.font.size = 13;
    text.lineFormat.visible = false;
  });
  await context.sync();
}

export const insertBusinessLayout: Tool = {
  name: "insert_business_layout",
  description: "Insert an editable PowerPoint business layout such as a timeline, process flow, comparison grid, phase plan, or estimate summary.",
  parameters: {
    type: "object",
    properties: {
      slideIndex: { type: "number", description: "Optional slide index. If omitted, the tool prefers creating a new slide from a matching template layout." },
      layoutType: { type: "string", enum: ["timeline", "phase_plan", "process_flow", "comparison_grid", "estimate_summary"] },
      title: { type: "string" },
      themeMode: { type: "string", enum: ["deck", "custom"] },
      items: {
        type: "array",
        items: {
          type: "object",
          properties: {
            title: { type: "string" },
            subtitle: { type: "string" },
            value: { type: "string" },
            colorToken: { type: "string" },
          },
          required: ["title"],
        },
      },
    },
    required: ["layoutType", "items"],
  },
  handler: async (args) => {
    const layoutArgs = resolvePowerPointTargetingArgs(args as InsertBusinessLayoutArgs);
    if (!layoutArgs.items?.length) {
      return toolFailure("items must contain at least one item.");
    }

    try {
      return await PowerPoint.run(async (context) => {
        let slide: PowerPoint.Slide;
        let slideIndex: number;

        if (Number.isInteger(layoutArgs.slideIndex) && (layoutArgs.slideIndex as number) >= 0) {
          slideIndex = layoutArgs.slideIndex as number;
          slide = await getSlideByIndex(context, slideIndex);
        } else {
          const templateLayout = await findTemplateLayout(context, layoutArgs.layoutType);
          if (templateLayout) {
            const created = await createSlideWithLayout(context, templateLayout);
            slide = created.slide;
            slideIndex = created.slideIndex;
            if (layoutArgs.title) {
              await applyTemplateBindingsToSlide(context, slide, [{ placeholderType: "Title", text: layoutArgs.title }]);
            }
          } else {
            context.presentation.slides.add();
            await context.sync();
            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();
            slideIndex = slides.items.length - 1;
            slide = slides.items[slideIndex];
          }
        }

        await drawBusinessLayout(context, slide, layoutArgs.layoutType, layoutArgs.items, layoutArgs.title);
        slide.load("id");
        await context.sync();
        return {
          resultType: "success",
          textResultForLlm: `Inserted ${layoutArgs.layoutType} layout on slide ${slideIndex + 1}.`,
          slideIndex,
          slideId: slide.id,
          toolTelemetry: { slideIndex, slideId: slide.id, layoutType: layoutArgs.layoutType, itemCount: layoutArgs.items.length },
        };
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
