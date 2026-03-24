import type { Tool } from "@github/copilot-sdk";

export const applyCellFormatting: Tool = {
  name: "apply_cell_formatting",
  description: `Apply formatting to cells in Excel.

Parameters:
- range: The cell range to format (e.g., "A1:D10", "B2", "A:A" for entire column)
- sheetName: Optional worksheet name. Defaults to active sheet.

Formatting options (all optional):
- bold: Make text bold
- italic: Make text italic
- underline: Underline text
- fontSize: Font size in points
- fontColor: Text color as hex (e.g., "FF0000" for red)
- backgroundColor: Cell fill color as hex (e.g., "FFFF00" for yellow)
- numberFormat: Excel number format (e.g., "$#,##0.00", "0%", "yyyy-mm-dd")
- horizontalAlignment: "left", "center", "right"
- borderStyle: "thin", "medium", "thick", "none"
- borderColor: Border color as hex

Examples:
- Bold headers: range="A1:E1", bold=true, backgroundColor="4472C4", fontColor="FFFFFF"
- Currency format: range="B2:B100", numberFormat="$#,##0.00"
- Percentage: range="C2:C50", numberFormat="0.0%"
- Center align: range="A1:Z1", horizontalAlignment="center"
- Add borders: range="A1:D10", borderStyle="thin", borderColor="000000"`,
  parameters: {
    type: "object",
    properties: {
      range: {
        type: "string",
        description: "The cell range to format (e.g., 'A1:D10').",
      },
      sheetName: {
        type: "string",
        description: "Optional worksheet name. Defaults to active sheet.",
      },
      bold: {
        type: "boolean",
        description: "Set text to bold.",
      },
      italic: {
        type: "boolean",
        description: "Set text to italic.",
      },
      underline: {
        type: "boolean",
        description: "Underline text.",
      },
      fontSize: {
        type: "number",
        description: "Font size in points.",
      },
      fontColor: {
        type: "string",
        description: "Text color as hex without # (e.g., 'FF0000').",
      },
      backgroundColor: {
        type: "string",
        description: "Cell fill color as hex without # (e.g., 'FFFF00').",
      },
      numberFormat: {
        type: "string",
        description: "Excel number format (e.g., '$#,##0.00', '0%').",
      },
      horizontalAlignment: {
        type: "string",
        enum: ["left", "center", "right"],
        description: "Horizontal text alignment.",
      },
      borderStyle: {
        type: "string",
        enum: ["thin", "medium", "thick", "none"],
        description: "Border line style.",
      },
      borderColor: {
        type: "string",
        description: "Border color as hex without # (e.g., '000000').",
      },
    },
    required: ["range"],
  },
  handler: async (args) => {
    const {
      range: rangeAddress,
      sheetName,
      bold,
      italic,
      underline,
      fontSize,
      fontColor,
      backgroundColor,
      numberFormat,
      horizontalAlignment,
      borderStyle,
      borderColor,
    } = args as {
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
      borderStyle?: string;
      borderColor?: string;
    };

    // Check if any formatting was specified
    const hasFormatting = bold !== undefined || italic !== undefined || underline !== undefined ||
      fontSize !== undefined || fontColor !== undefined || backgroundColor !== undefined ||
      numberFormat !== undefined || horizontalAlignment !== undefined ||
      borderStyle !== undefined;

    if (!hasFormatting) {
      return { textResultForLlm: "No formatting options specified.", resultType: "failure", error: "No formatting", toolTelemetry: {} };
    }

    try {
      return await Excel.run(async (context) => {
        // Get the target worksheet
        let sheet: Excel.Worksheet;
        if (sheetName) {
          sheet = context.workbook.worksheets.getItem(sheetName);
        } else {
          sheet = context.workbook.worksheets.getActiveWorksheet();
        }
        sheet.load("name");

        // Get the range
        const range = sheet.getRange(rangeAddress);
        range.load("address");
        await context.sync();

        const format = range.format;
        const font = format.font;

        // Apply font formatting
        if (bold !== undefined) font.bold = bold;
        if (italic !== undefined) font.italic = italic;
        if (underline !== undefined) font.underline = underline ? Excel.RangeUnderlineStyle.single : Excel.RangeUnderlineStyle.none;
        if (fontSize !== undefined) font.size = fontSize;
        if (fontColor !== undefined) font.color = fontColor.startsWith("#") ? fontColor : `#${fontColor}`;

        // Apply fill
        if (backgroundColor !== undefined) {
          format.fill.color = backgroundColor.startsWith("#") ? backgroundColor : `#${backgroundColor}`;
        }

        // Apply number format
        if (numberFormat !== undefined) {
          range.numberFormat = [[numberFormat]];
        }

        // Apply alignment
        if (horizontalAlignment !== undefined) {
          const alignmentMap: { [key: string]: Excel.HorizontalAlignment } = {
            "left": Excel.HorizontalAlignment.left,
            "center": Excel.HorizontalAlignment.center,
            "right": Excel.HorizontalAlignment.right,
          };
          format.horizontalAlignment = alignmentMap[horizontalAlignment] || Excel.HorizontalAlignment.general;
        }

        // Apply borders
        if (borderStyle !== undefined) {
          const styleMap: { [key: string]: Excel.BorderLineStyle } = {
            "thin": Excel.BorderLineStyle.thin,
            "medium": Excel.BorderLineStyle.medium,
            "thick": Excel.BorderLineStyle.thick,
            "none": Excel.BorderLineStyle.none,
          };
          const lineStyle = styleMap[borderStyle] || Excel.BorderLineStyle.thin;
          const color = borderColor ? (borderColor.startsWith("#") ? borderColor : `#${borderColor}`) : "#000000";

          const borders = format.borders;
          const borderTypes = [
            Excel.BorderIndex.edgeTop,
            Excel.BorderIndex.edgeBottom,
            Excel.BorderIndex.edgeLeft,
            Excel.BorderIndex.edgeRight,
          ];

          for (const borderType of borderTypes) {
            const border = borders.getItem(borderType);
            border.style = lineStyle;
            if (lineStyle !== Excel.BorderLineStyle.none) {
              border.color = color;
            }
          }
        }

        await context.sync();

        // Build confirmation message
        const applied: string[] = [];
        if (bold !== undefined) applied.push(bold ? "bold" : "not bold");
        if (italic !== undefined) applied.push(italic ? "italic" : "not italic");
        if (underline !== undefined) applied.push(underline ? "underlined" : "not underlined");
        if (fontSize !== undefined) applied.push(`${fontSize}pt font`);
        if (fontColor !== undefined) applied.push(`font color #${fontColor.replace("#", "")}`);
        if (backgroundColor !== undefined) applied.push(`fill #${backgroundColor.replace("#", "")}`);
        if (numberFormat !== undefined) applied.push(`format "${numberFormat}"`);
        if (horizontalAlignment !== undefined) applied.push(`${horizontalAlignment} aligned`);
        if (borderStyle !== undefined) applied.push(`${borderStyle} borders`);

        return `Applied formatting to ${range.address} in "${sheet.name}": ${applied.join(", ")}.`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
