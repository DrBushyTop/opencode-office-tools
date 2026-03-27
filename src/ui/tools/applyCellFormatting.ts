import type { Tool } from "./types";
import { getWorksheet, normalizeExcelColor, toolFailure } from "./excelShared";

export const applyCellFormatting: Tool = {
  name: "apply_cell_formatting",
  description: "Apply formatting to Excel cells, including fonts, fills, number formats, alignment, wrapping, merging, borders, and row or column sizing.",
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
      merge: { type: "boolean", description: "Merge or unmerge the target range." },
      mergeAcross: { type: "boolean", description: "When merging, merge each row separately instead of the full range." },
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
    const update = args as {
      range: string;
      sheetName?: string;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      fontSize?: number;
      fontColor?: string;
      backgroundColor?: string;
      numberFormat?: string;
      horizontalAlignment?: string;
      verticalAlignment?: string;
      wrapText?: boolean;
      merge?: boolean;
      mergeAcross?: boolean;
      borderStyle?: string;
      borderColor?: string;
      interiorBorders?: boolean;
      rowHeight?: number;
      columnWidth?: number;
      autoFitRows?: boolean;
      autoFitColumns?: boolean;
    };

    const hasFormatting = Object.entries(update).some(([key, value]) => key !== "range" && key !== "sheetName" && value !== undefined);
    if (!hasFormatting) {
      return toolFailure("No formatting options specified.");
    }
    if (update.rowHeight !== undefined && update.rowHeight < 0) {
      return toolFailure("rowHeight must be non-negative.");
    }
    if (update.columnWidth !== undefined && update.columnWidth < 0) {
      return toolFailure("columnWidth must be non-negative.");
    }

    try {
      return await Excel.run(async (context) => {
        const sheet = await getWorksheet(context, update.sheetName);
        const range = sheet.getRange(update.range);
        range.load(["address", "rowCount", "columnCount"]);
        await context.sync();

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
        if (update.numberFormat !== undefined) {
          range.numberFormat = Array.from({ length: range.rowCount }, () => Array.from({ length: range.columnCount }, () => update.numberFormat));
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

        if (update.wrapText !== undefined) format.wrapText = update.wrapText;
        if (update.rowHeight !== undefined) format.rowHeight = update.rowHeight;
        if (update.columnWidth !== undefined) format.columnWidth = update.columnWidth;
        if (update.autoFitRows) format.autofitRows();
        if (update.autoFitColumns) format.autofitColumns();
        if (update.merge !== undefined) {
          if (update.merge) {
            range.merge(Boolean(update.mergeAcross));
          } else {
            range.unmerge();
          }
        }

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
        if (update.merge !== undefined) applied.push(update.merge ? "merged" : "unmerged");
        if (update.borderStyle !== undefined) applied.push(`${update.borderStyle} borders${update.interiorBorders ? " including interior" : ""}`);
        if (update.rowHeight !== undefined) applied.push(`row height ${update.rowHeight}`);
        if (update.columnWidth !== undefined) applied.push(`column width ${update.columnWidth}`);
        if (update.autoFitRows) applied.push("auto-fit rows");
        if (update.autoFitColumns) applied.push("auto-fit columns");

        return `Applied formatting to ${range.address} in ${sheet.name}: ${applied.join(", ")}.`;
      });
    } catch (error: unknown) {
      return toolFailure(error);
    }
  },
};
