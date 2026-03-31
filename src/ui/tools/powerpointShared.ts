import type { ToolResultFailure } from "./types";
import { createToolFailure, describeErrorWithCode, summarizePlainText as summarizeSharedPlainText } from "./toolShared";

export function toolFailure(error: unknown, hint?: string): ToolResultFailure {
  return createToolFailure(error, { hint, describe: describeErrorWithCode });
}

/** Extract a descriptive message from an error, including the Office.js error code when available. */
export function describeError(error: unknown): string {
  return describeErrorWithCode(error);
}

export function formatAvailableSlideIndexes(slideCount: number) {
  if (slideCount <= 0) return "Presentation has no slides.";

  const preview = Array.from({ length: Math.min(slideCount, 8) }, (_, index) => String(index)).join(", ");
  return slideCount <= 8
    ? `Available slideIndex values: ${preview}.`
    : `Available slideIndex values: ${preview}, ... ${slideCount - 1}.`;
}

export function invalidSlideIndexMessage(slideIndex: number, slideCount: number) {
  return `Invalid slideIndex ${slideIndex}. ${formatAvailableSlideIndexes(slideCount)}`;
}

export function formatAvailableShapeTargets(
  slideIndex: number,
  shapes: Array<{ id?: string | null; name?: string | null }>,
) {
  if (shapes.length === 0) {
    return `Slide ${slideIndex + 1} has no shapes.`;
  }

  const preview = shapes
    .slice(0, 6)
    .map((shape, index) => {
      const id = readOfficeValue(() => shape.id, "(missing)");
      const name = readOfficeValue(() => shape.name, "");
      return `shapeIndex ${index}: id=${id || "(missing)"}, name=${JSON.stringify(name || "")}`;
    })
    .join("; ");
  const suffix = shapes.length > 6 ? "; ..." : "";

  return `Available shapes on slide ${slideIndex + 1}: ${preview}${suffix}`;
}

export function roundTripRefreshHint() {
  return "Hint: this can happen after an Open XML round-trip replaces a slide. Re-run get_presentation_overview to refresh current slideIndex values, then retry with current shape refs or shapeId values from the latest tool results.";
}

export function roundTripSlideRefreshHint() {
  return "Hint: this can happen after an Open XML round-trip replaces a slide. Re-run get_presentation_overview to refresh current slideIndex values, then retry.";
}

export function shouldAddRoundTripRefreshHint(error: unknown) {
  const text = describeError(error);
  return /object can not be found here|object cannot be found here|invalidobjectpath|invalidargument|item.*does not exist/i.test(text);
}

export function shouldAddRoundTripShapeTargetRefreshHint(error: unknown) {
  const text = describeError(error);
  return shouldAddRoundTripRefreshHint(error)
    || /^Shape .+ was not found on slide \d+\./i.test(text)
    || /^Slide .+ was not found in the current presentation\./i.test(text)
    || /^Could not find shape ref .+ on exported slide .+/i.test(text);
}

export function isPowerPointRequirementSetSupported(version: string) {
  return Office.context.requirements.isSetSupported("PowerPointApi", version);
}

export const supportsPowerPointPlaceholders = () => isPowerPointRequirementSetSupported("1.8");

export function summarizePlainText(text: string, limit = 100) {
  return summarizeSharedPlainText(text, limit);
}

const themeColorKeys = [
  "Dark1",
  "Light1",
  "Dark2",
  "Light2",
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
  "Hyperlink",
  "FollowedHyperlink",
] as const;

export type PowerPointThemeColorKey = (typeof themeColorKeys)[number];

export interface PowerPointThemeColors {
  Dark1?: string;
  Light1?: string;
  Dark2?: string;
  Light2?: string;
  Accent1?: string;
  Accent2?: string;
  Accent3?: string;
  Accent4?: string;
  Accent5?: string;
  Accent6?: string;
  Hyperlink?: string;
  FollowedHyperlink?: string;
}

export async function loadThemeColors(context: PowerPoint.RequestContext, master: PowerPoint.SlideMaster): Promise<PowerPointThemeColors> {
  const requests = themeColorKeys.map((color) => ({ color, value: master.themeColorScheme.getThemeColor(color) }));
  await context.sync();
  return Object.fromEntries(requests.map((entry) => [entry.color, parseColor(entry.value.value)])) as PowerPointThemeColors;
}

export function pickThemeColor(colors: PowerPointThemeColors | null | undefined, key: PowerPointThemeColorKey, fallback: string) {
  return colors?.[key] || fallback;
}

export function normalizeHexColor(value: string) {
  const trimmed = value.trim();
  if (trimmed.startsWith("#")) return trimmed;
  return /^[0-9A-Fa-f]{6}$/.test(trimmed) ? `#${trimmed}` : trimmed;
}


export function parseColor(value: string | null | undefined) {
  if (!value) return "(none)";
  if (value.startsWith("#")) return value;
  return /^[0-9A-Fa-f]{6}$/.test(value) ? `#${value}` : value;
}

export function readOfficeValue<T>(read: () => T, fallback: T): T {
  try {
    const value = read();
    return value ?? fallback;
  } catch {
    return fallback;
  }
}

export interface PowerPointShapeSummary {
  index: number;
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  rotation?: number;
  zOrderPosition?: number;
  visible?: boolean;
  text?: string;
  placeholderType?: string;
  placeholderContainedType?: string | null;
  altTextTitle?: string;
  altTextDescription?: string;
  fillColor?: string;
  fillType?: string;
  lineColor?: string;
  lineWeight?: number;
  lineDashStyle?: string;
  tableInfo?: { rowCount: number; columnCount: number; values?: string[][] };
}

export async function loadShapeSummaries(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
  options: { includeText?: boolean; includeFormatting?: boolean; includeTableValues?: boolean } = {},
) {
  const { includeText = true, includeFormatting = true, includeTableValues = true } = options;
  const includePlaceholders = supportsPowerPointPlaceholders();

  for (const shape of shapes) {
    shape.load(["id", "name", "type", "left", "top", "width", "height", "rotation", "zOrderPosition", "visible", "altTextTitle", "altTextDescription"]);
  }
  await context.sync();

  const placeholderFormats = includePlaceholders
    ? shapes.map((shape) => readOfficeValue(
      () => (shape.type === PowerPoint.ShapeType.placeholder ? shape.placeholderFormat : null),
      null,
    ))
    : shapes.map(() => null);
  if (includePlaceholders) {
    for (const placeholder of placeholderFormats) {
      placeholder?.load(["type", "containedType"]);
    }
  }

  const textFrames = includeText ? shapes.map((shape) => shape.getTextFrameOrNullObject()) : [];
  for (const frame of textFrames) {
    frame.load(["isNullObject", "hasText"]);
  }

  const fills = includeFormatting ? shapes.map((shape) => shape.fill) : [];
  const lines = includeFormatting ? shapes.map((shape) => shape.lineFormat) : [];
  for (const fill of fills) {
    fill.load(["foregroundColor", "transparency", "type"]);
  }
  for (const line of lines) {
    line.load(["color", "weight", "dashStyle", "visible"]);
  }

  const tables = shapes.map((shape) => readOfficeValue(
    () => (shape.type === PowerPoint.ShapeType.table ? shape.getTable() : null),
    null,
  ));
  for (const table of tables) {
    table?.load(includeTableValues ? ["rowCount", "columnCount", "values"] : ["rowCount", "columnCount"]);
  }
  await context.sync();

  for (const frame of textFrames) {
    if (!frame.isNullObject && frame.hasText) {
      frame.textRange.load("text");
    }
  }
  await context.sync();

  return shapes.map<PowerPointShapeSummary>((shape, index) => ({
    index,
    id: readOfficeValue(() => shape.id, `(missing ${index})`),
    name: readOfficeValue(() => shape.name, "") || `(unnamed ${index})`,
    type: readOfficeValue(() => String(shape.type), "Unknown"),
    left: readOfficeValue(() => shape.left, 0),
    top: readOfficeValue(() => shape.top, 0),
    width: readOfficeValue(() => shape.width, 0),
    height: readOfficeValue(() => shape.height, 0),
    rotation: readOfficeValue(() => shape.rotation, undefined),
    zOrderPosition: readOfficeValue(() => shape.zOrderPosition, undefined),
    visible: readOfficeValue(() => shape.visible, undefined),
    text: includeText && textFrames[index] && !textFrames[index].isNullObject && textFrames[index].hasText
      ? readOfficeValue(() => textFrames[index].textRange.text, "")
      : includeText && textFrames[index] && !textFrames[index].isNullObject
        ? ""
        : undefined,
    placeholderType: readOfficeValue(() => (placeholderFormats[index]?.type ? String(placeholderFormats[index]?.type) : undefined), undefined),
    placeholderContainedType: readOfficeValue(
      () => (placeholderFormats[index]?.containedType ? String(placeholderFormats[index]?.containedType) : placeholderFormats[index] ? null : undefined),
      undefined,
    ),
    altTextTitle: readOfficeValue(() => shape.altTextTitle, ""),
    altTextDescription: readOfficeValue(() => shape.altTextDescription, ""),
    fillColor: includeFormatting ? parseColor(readOfficeValue(() => fills[index]?.foregroundColor, undefined)) : undefined,
    fillType: includeFormatting ? readOfficeValue(() => (fills[index]?.type ? String(fills[index]?.type) : undefined), undefined) : undefined,
    lineColor: includeFormatting ? parseColor(readOfficeValue(() => lines[index]?.color, undefined)) : undefined,
    lineWeight: includeFormatting ? readOfficeValue(() => lines[index]?.weight, undefined) : undefined,
    lineDashStyle: includeFormatting ? readOfficeValue(() => (lines[index]?.dashStyle ? String(lines[index]?.dashStyle) : undefined), undefined) : undefined,
    tableInfo: tables[index]
      ? {
          rowCount: readOfficeValue(() => tables[index]!.rowCount, 0),
          columnCount: readOfficeValue(() => tables[index]!.columnCount, 0),
          values: includeTableValues ? readOfficeValue(() => tables[index]!.values, undefined) : undefined,
        }
      : undefined,
  }));
}

export function formatShapeSummary(shape: PowerPointShapeSummary, detail = false) {
  const lines = [
    `- Shape ${shape.index}: id=${shape.id}, name=${JSON.stringify(shape.name)}, type=${shape.type}, box=(${shape.left}, ${shape.top}, ${shape.width}, ${shape.height})`,
  ];

  if (shape.placeholderType) {
    lines.push(`  placeholder=${shape.placeholderType}${shape.placeholderContainedType ? `, contained=${shape.placeholderContainedType}` : ""}`);
  }
  if (shape.text !== undefined) {
    lines.push(`  text=${detail ? JSON.stringify(shape.text || "") : summarizePlainText(shape.text || "")}`);
  }
  if (shape.tableInfo) {
    lines.push(`  table=${shape.tableInfo.rowCount}x${shape.tableInfo.columnCount}${detail && shape.tableInfo.values ? `, values=${JSON.stringify(shape.tableInfo.values)}` : ""}`);
  }
  if (shape.altTextTitle || shape.altTextDescription) {
    lines.push(`  altText=${JSON.stringify(shape.altTextTitle || shape.altTextDescription || "")}`);
  }
  if (shape.fillType || shape.lineColor) {
    lines.push(`  style=fill ${shape.fillType || "unknown"} ${shape.fillColor || ""}, line ${shape.lineColor || "(none)"} ${shape.lineWeight ?? ""}`.trim());
  }

  return lines.join("\n");
}
