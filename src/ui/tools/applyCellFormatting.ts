import { z } from "zod";
import type { Tool } from "./types";
import { getWorksheet, nonNegativeFiniteNumberSchema, normalizeExcelColor, parseToolArgs } from "./excelShared";
import { createToolFailure, describeErrorWithCode } from "./toolShared";

const horizontalAlignmentSchema = z.enum(["left", "center", "right", "general", "fill", "justify", "centerAcrossSelection", "distributed"]);
const verticalAlignmentSchema = z.enum(["top", "center", "bottom", "justify", "distributed"]);
const borderStyleSchema = z.enum(["thin", "medium", "thick", "none", "double", "dashed", "dotted"]);

function isBlankishString(value: string | undefined) {
  return value !== undefined && value.trim() === "";
}

function looksUnboundedRange(rangeAddress: string) {
  return /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/.test(rangeAddress.trim()) || /^\$?\d+:\$?\d+$/.test(rangeAddress.trim());
}

function buildRepeated2DArray(rowCount: number, columnCount: number, value: string) {
  return Array.from({ length: rowCount }, () => Array.from({ length: columnCount }, () => value));
}

function normalizeFormattingUpdate(update: z.infer<typeof applyCellFormattingArgsSchema>) {
  const normalized = { ...update };

  if (isBlankishString(normalized.fontColor)) normalized.fontColor = undefined;
  if (isBlankishString(normalized.backgroundColor)) normalized.backgroundColor = undefined;
  if (isBlankishString(normalized.numberFormat)) normalized.numberFormat = undefined;
  if (isBlankishString(normalized.borderColor)) normalized.borderColor = undefined;

  // Model-generated default payloads often use 0 as a placeholder for sizing.
  // Treat that as "omit" instead of sending an invalid or expensive update.
  if (normalized.fontSize === 0) normalized.fontSize = undefined;
  if (normalized.rowHeight === 0) normalized.rowHeight = undefined;
  if (normalized.columnWidth === 0) normalized.columnWidth = undefined;

  // Omit explicit false for merge when paired with a model-style default payload.
  // Real unmerge intent can still be expressed by sending merge=false without the
  // placeholder reset pattern.
  if (
    normalized.merge === false
    && normalized.mergeAcross === false
    && normalized.bold === false
    && normalized.italic === false
    && normalized.underline === false
    && normalized.wrapText === false
    && normalized.autoFitRows === false
    && normalized.autoFitColumns === false
  ) {
    normalized.merge = undefined;
    normalized.mergeAcross = undefined;
  }

  return normalized;
}

const applyCellFormattingArgsSchema = z.object({
  range: z.string(),
  sheetName: z.string().optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  fontSize: z.number().optional(),
  fontColor: z.string().optional(),
  backgroundColor: z.string().optional(),
  numberFormat: z.string().optional(),
  horizontalAlignment: horizontalAlignmentSchema.optional(),
  verticalAlignment: verticalAlignmentSchema.optional(),
  wrapText: z.boolean().optional(),
  merge: z.boolean().optional(),
  mergeAcross: z.boolean().optional(),
  borderStyle: borderStyleSchema.optional(),
  borderColor: z.string().optional(),
  interiorBorders: z.boolean().optional(),
  rowHeight: nonNegativeFiniteNumberSchema("rowHeight must be non-negative.").optional(),
  columnWidth: nonNegativeFiniteNumberSchema("columnWidth must be non-negative.").optional(),
  autoFitRows: z.boolean().optional(),
  autoFitColumns: z.boolean().optional(),
});

export const applyCellFormatting: Tool = {
  name: "apply_cell_formatting",
  description: "Apply formatting to Excel cells, including fonts, fills, number formats, alignment, wrapping, borders, row or column sizing, and optional merge-state changes. Pass only the fields you want to change; omit unchanged/default values. Omit merge to leave merge state unchanged.",
  parameters: {
    type: "object",
    properties: {
      range: { type: "string", description: "Target range such as 'A1:D10'." },
      sheetName: { type: "string", description: "Optional worksheet name. Defaults to the active sheet." },
      bold: { type: "boolean" },
      italic: { type: "boolean" },
      underline: { type: "boolean" },
      fontSize: { type: "number" },
      fontColor: { type: "string" },
      backgroundColor: { type: "string" },
      numberFormat: { type: "string" },
      horizontalAlignment: { type: "string", enum: ["left", "center", "right", "general", "fill", "justify", "centerAcrossSelection", "distributed"] },
      verticalAlignment: { type: "string", enum: ["top", "center", "bottom", "justify", "distributed"] },
      wrapText: { type: "boolean" },
      merge: { type: "boolean", description: "Set to true to merge the target range, or false to actively unmerge it. Omit this field to leave merge state unchanged. Excel table cells cannot be merged." },
      mergeAcross: { type: "boolean", description: "When merge=true, merge each row separately instead of the full range." },
      borderStyle: { type: "string", enum: ["thin", "medium", "thick", "none", "double", "dashed", "dotted"] },
      borderColor: { type: "string" },
      interiorBorders: { type: "boolean", description: "Also apply border formatting to inside horizontal and vertical borders." },
      rowHeight: { type: "number" },
      columnWidth: { type: "number" },
      autoFitRows: { type: "boolean" },
      autoFitColumns: { type: "boolean" },
    },
    required: ["range"],
  },
  handler: async (args) => {
    const parsedArgs = parseToolArgs(applyCellFormattingArgsSchema, args);
    if (!parsedArgs.success) return parsedArgs.failure;

    const update = normalizeFormattingUpdate(parsedArgs.data);

    const hasFormatting = Object.entries(update).some(([key, value]) => key !== "range" && key !== "sheetName" && value !== undefined);
    if (!hasFormatting) {
      return createToolFailure("No formatting options specified.");
    }

    try {
      return await Excel.run(async (context) => {
        const sheet = await getWorksheet(context, update.sheetName);
        const targetRange = sheet.getRange(update.range);
        const isUnbounded = looksUnboundedRange(update.range);
        const range = isUnbounded && typeof (targetRange as Excel.Range & { getUsedRangeOrNullObject?: (valuesOnly?: boolean) => Excel.Range }).getUsedRangeOrNullObject === "function"
          ? (targetRange as Excel.Range & { getUsedRangeOrNullObject: (valuesOnly?: boolean) => Excel.Range }).getUsedRangeOrNullObject(true)
          : targetRange;
        range.load(["address", "rowCount", "columnCount", "isNullObject"]);
        await context.sync();

        if ((range as Excel.Range & { isNullObject?: boolean }).isNullObject) {
          return createToolFailure(`Range ${update.range} in ${sheet.name} has no used cells to format.`);
        }

        if (isUnbounded && update.numberFormat !== undefined) {
          // Narrow unbounded number-format writes to used cells to avoid huge
          // 2D payloads like full-column A:A across 1M rows.
          range.numberFormat = buildRepeated2DArray(range.rowCount, range.columnCount, update.numberFormat);
        }

        let overlapsTable = false;
        if (update.merge !== undefined && typeof (range as Excel.Range & { getTables?: (fullyContained?: boolean) => Excel.TableScopedCollection }).getTables === "function") {
          const tables = (range as Excel.Range & { getTables: (fullyContained?: boolean) => Excel.TableScopedCollection }).getTables(false);
          tables.load("items/name");
          await context.sync();
          overlapsTable = tables.items.length > 0;

          if (update.merge) {
            return createToolFailure(`Cannot merge ${range.address} because it overlaps an Excel table. Convert the table to a normal range first or omit merge.`);
          }
        }

        const format = range.format;
        const font = format.font;
        if (update.bold !== undefined) font.bold = update.bold;
        if (update.italic !== undefined) font.italic = update.italic;
        if (update.underline !== undefined) {
          font.underline = update.underline ? Excel.RangeUnderlineStyle.single : Excel.RangeUnderlineStyle.none;
        }
        if (update.fontSize !== undefined) font.size = update.fontSize;
        if (update.fontColor !== undefined) font.color = normalizeExcelColor(update.fontColor);
        if (update.backgroundColor !== undefined) format.fill.color = normalizeExcelColor(update.backgroundColor);
        if (!isUnbounded && update.numberFormat !== undefined) {
          range.numberFormat = buildRepeated2DArray(range.rowCount, range.columnCount, update.numberFormat);
        }

        if (update.horizontalAlignment !== undefined) {
          const alignmentMap: Record<string, Excel.HorizontalAlignment> = {
            left: Excel.HorizontalAlignment.left,
            center: Excel.HorizontalAlignment.center,
            right: Excel.HorizontalAlignment.right,
            general: Excel.HorizontalAlignment.general,
            fill: Excel.HorizontalAlignment.fill,
            justify: Excel.HorizontalAlignment.justify,
            centerAcrossSelection: Excel.HorizontalAlignment.centerAcrossSelection,
            distributed: Excel.HorizontalAlignment.distributed,
          };
          format.horizontalAlignment = alignmentMap[update.horizontalAlignment] || Excel.HorizontalAlignment.general;
        }

        if (update.verticalAlignment !== undefined) {
          const verticalMap: Record<string, Excel.VerticalAlignment> = {
            top: Excel.VerticalAlignment.top,
            center: Excel.VerticalAlignment.center,
            bottom: Excel.VerticalAlignment.bottom,
            justify: Excel.VerticalAlignment.justify,
            distributed: Excel.VerticalAlignment.distributed,
          };
          format.verticalAlignment = verticalMap[update.verticalAlignment] || Excel.VerticalAlignment.bottom;
        }

        // Apply merge state changes before sizing operations. Excel can reject
        // row-height and autofit updates against merged ranges.
        let mergeSummary: string | null = null;
        if (update.merge !== undefined) {
          if (update.merge) {
            range.merge(Boolean(update.mergeAcross));
            mergeSummary = "merged";
          } else if (!overlapsTable) {
            range.unmerge();
            mergeSummary = "unmerged";
          } else {
            mergeSummary = "merge unchanged (table cells cannot be merged or unmerged)";
          }
        }

        if (update.wrapText !== undefined) format.wrapText = update.wrapText;
        if (update.rowHeight !== undefined) format.rowHeight = update.rowHeight;
        if (update.columnWidth !== undefined) format.columnWidth = update.columnWidth;
        if (update.autoFitRows) format.autofitRows();
        if (update.autoFitColumns) format.autofitColumns();

        if (update.borderStyle !== undefined) {
          const styleMap: Record<string, Excel.BorderLineStyle> = {
            thin: Excel.BorderLineStyle.continuous,
            medium: Excel.BorderLineStyle.continuous,
            thick: Excel.BorderLineStyle.continuous,
            none: Excel.BorderLineStyle.none,
            double: Excel.BorderLineStyle.double,
            dashed: Excel.BorderLineStyle.dash,
            dotted: Excel.BorderLineStyle.dot,
          };
          const weightMap: Record<string, Excel.BorderWeight> = {
            thin: Excel.BorderWeight.thin,
            medium: Excel.BorderWeight.medium,
            thick: Excel.BorderWeight.thick,
            none: Excel.BorderWeight.thin,
            double: Excel.BorderWeight.thin,
            dashed: Excel.BorderWeight.thin,
            dotted: Excel.BorderWeight.thin,
          };
          const lineStyle = styleMap[update.borderStyle] || Excel.BorderLineStyle.continuous;
          const weight = weightMap[update.borderStyle] || Excel.BorderWeight.thin;
          const color = update.borderColor ? normalizeExcelColor(update.borderColor) : "#000000";
          const borderTypes = [
            Excel.BorderIndex.edgeTop,
            Excel.BorderIndex.edgeBottom,
            Excel.BorderIndex.edgeLeft,
            Excel.BorderIndex.edgeRight,
            ...(update.interiorBorders ? [Excel.BorderIndex.insideHorizontal, Excel.BorderIndex.insideVertical] : []),
          ];

          for (const borderType of borderTypes) {
            const border = format.borders.getItem(borderType);
            border.style = lineStyle;
            if (lineStyle !== Excel.BorderLineStyle.none) {
              border.color = color;
              border.weight = weight;
            }
          }
        }

        await context.sync();

        const applied: string[] = [];
        if (update.bold !== undefined) applied.push(update.bold ? "bold" : "not bold");
        if (update.italic !== undefined) applied.push(update.italic ? "italic" : "not italic");
        if (update.underline !== undefined) applied.push(update.underline ? "underlined" : "not underlined");
        if (update.fontSize !== undefined) applied.push(`${update.fontSize}pt font`);
        if (update.fontColor !== undefined) applied.push(`font ${normalizeExcelColor(update.fontColor)}`);
        if (update.backgroundColor !== undefined) applied.push(`fill ${normalizeExcelColor(update.backgroundColor)}`);
        if (update.numberFormat !== undefined) applied.push(`format ${JSON.stringify(update.numberFormat)}`);
        if (update.horizontalAlignment !== undefined) applied.push(`${update.horizontalAlignment} horizontal alignment`);
        if (update.verticalAlignment !== undefined) applied.push(`${update.verticalAlignment} vertical alignment`);
        if (update.wrapText !== undefined) applied.push(update.wrapText ? "wrap text on" : "wrap text off");
        if (mergeSummary) applied.push(mergeSummary);
        if (update.borderStyle !== undefined) applied.push(`${update.borderStyle} borders${update.interiorBorders ? " including interior" : ""}`);
        if (update.rowHeight !== undefined) applied.push(`row height ${update.rowHeight}`);
        if (update.columnWidth !== undefined) applied.push(`column width ${update.columnWidth}`);
        if (update.autoFitRows) applied.push("auto-fit rows");
        if (update.autoFitColumns) applied.push("auto-fit columns");

        return `Applied formatting to ${range.address} in ${sheet.name}: ${applied.join(", ")}.`;
      });
    } catch (error: unknown) {
      return createToolFailure(error, { describe: describeErrorWithCode });
    }
  },
};
