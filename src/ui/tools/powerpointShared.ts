import type { ToolResultFailure } from "./types";

export function toolFailure(error: unknown): ToolResultFailure {
  const message = error instanceof Error ? error.message : String(error);
  return { textResultForLlm: message, resultType: "failure", error: message, toolTelemetry: {} };
}

export function isPowerPointRequirementSetSupported(version: string) {
  return Office.context.requirements.isSetSupported("PowerPointApi", version);
}

export const supportsPowerPointPlaceholders = () => isPowerPointRequirementSetSupported("1.8");

export function summarizePlainText(text: string, limit = 100) {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  if (!normalized) return "(empty)";
  return normalized.length > limit ? `${normalized.slice(0, limit - 3)}...` : normalized;
}

export function parseColor(value: string | null | undefined) {
  if (!value) return "(none)";
  if (value.startsWith("#")) return value;
  return /^[0-9A-Fa-f]{6}$/.test(value) ? `#${value}` : value;
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
    ? shapes.map((shape) => shape.type === PowerPoint.ShapeType.placeholder ? shape.placeholderFormat : null)
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

  const tables = shapes.map((shape) => shape.type === PowerPoint.ShapeType.table ? shape.getTable() : null);
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
    id: shape.id,
    name: shape.name || `(unnamed ${index})`,
    type: String(shape.type),
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
    rotation: shape.rotation,
    zOrderPosition: shape.zOrderPosition,
    visible: shape.visible,
    text: includeText && textFrames[index] && !textFrames[index].isNullObject && textFrames[index].hasText ? textFrames[index].textRange.text || "" : undefined,
    placeholderType: placeholderFormats[index]?.type ? String(placeholderFormats[index]?.type) : undefined,
    placeholderContainedType: placeholderFormats[index]?.containedType ? String(placeholderFormats[index]?.containedType) : placeholderFormats[index] ? null : undefined,
    altTextTitle: shape.altTextTitle || "",
    altTextDescription: shape.altTextDescription || "",
    fillColor: includeFormatting ? parseColor(fills[index]?.foregroundColor) : undefined,
    fillType: includeFormatting && fills[index]?.type ? String(fills[index]?.type) : undefined,
    lineColor: includeFormatting ? parseColor(lines[index]?.color) : undefined,
    lineWeight: includeFormatting ? lines[index]?.weight : undefined,
    lineDashStyle: includeFormatting && lines[index]?.dashStyle ? String(lines[index]?.dashStyle) : undefined,
    tableInfo: tables[index]
      ? {
          rowCount: tables[index]!.rowCount,
          columnCount: tables[index]!.columnCount,
          values: includeTableValues ? tables[index]!.values : undefined,
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
