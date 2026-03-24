import type { Tool } from "@github/copilot-sdk";

export const applyStyleToSelection: Tool = {
  name: "apply_style_to_selection",
  description: `Apply formatting styles to the currently selected text in Word.

All parameters are optional - only specified styles will be applied.

Parameters:
- bold: Set text to bold (true) or remove bold (false)
- italic: Set text to italic (true) or remove italic (false)
- underline: Set text to underline (true) or remove underline (false)
- strikethrough: Set strikethrough (true) or remove it (false)
- fontSize: Font size in points (e.g., 12, 14, 24)
- fontName: Font family name (e.g., "Arial", "Times New Roman", "Calibri")
- fontColor: Text color as hex string (e.g., "FF0000" for red, "0000FF" for blue)
- highlightColor: Highlight/background color. Use Word highlight colors: "yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "gray25", "gray50", "black", or "noHighlight" to remove

Examples:
- Make text bold and red: bold=true, fontColor="FF0000"
- Increase font size: fontSize=16
- Highlight important text: highlightColor="yellow"
- Apply multiple styles: bold=true, italic=true, fontSize=14, fontName="Arial"`,
  parameters: {
    type: "object",
    properties: {
      bold: {
        type: "boolean",
        description: "Set to true for bold, false to remove bold.",
      },
      italic: {
        type: "boolean",
        description: "Set to true for italic, false to remove italic.",
      },
      underline: {
        type: "boolean",
        description: "Set to true for underline, false to remove underline.",
      },
      strikethrough: {
        type: "boolean",
        description: "Set to true for strikethrough, false to remove it.",
      },
      fontSize: {
        type: "number",
        description: "Font size in points.",
      },
      fontName: {
        type: "string",
        description: "Font family name (e.g., 'Arial', 'Calibri').",
      },
      fontColor: {
        type: "string",
        description: "Text color as hex string without # (e.g., 'FF0000' for red).",
      },
      highlightColor: {
        type: "string",
        description: "Highlight color name (e.g., 'yellow', 'green', 'noHighlight').",
      },
    },
    required: [],
  },
  handler: async (args) => {
    const {
      bold,
      italic,
      underline,
      strikethrough,
      fontSize,
      fontName,
      fontColor,
      highlightColor,
    } = args as {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      strikethrough?: boolean;
      fontSize?: number;
      fontName?: string;
      fontColor?: string;
      highlightColor?: string;
    };

    // Check if any style was specified
    const hasStyles = bold !== undefined || italic !== undefined || underline !== undefined ||
      strikethrough !== undefined || fontSize !== undefined || fontName !== undefined ||
      fontColor !== undefined || highlightColor !== undefined;

    if (!hasStyles) {
      return { textResultForLlm: "No styles specified. Provide at least one style parameter.", resultType: "failure", error: "No styles", toolTelemetry: {} };
    }

    try {
      return await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim().length === 0) {
          return "No text selected. Please select some text first.";
        }

        const font = selection.font;

        // Apply each specified style
        if (bold !== undefined) {
          font.bold = bold;
        }
        if (italic !== undefined) {
          font.italic = italic;
        }
        if (underline !== undefined) {
          font.underline = underline ? Word.UnderlineType.single : Word.UnderlineType.none;
        }
        if (strikethrough !== undefined) {
          font.strikeThrough = strikethrough;
        }
        if (fontSize !== undefined) {
          font.size = fontSize;
        }
        if (fontName !== undefined) {
          font.name = fontName;
        }
        if (fontColor !== undefined) {
          font.color = fontColor.startsWith("#") ? fontColor : `#${fontColor}`;
        }
        if (highlightColor !== undefined) {
          // Map string to Word.HighlightColor
          const colorMap: { [key: string]: Word.HighlightColor } = {
            "yellow": Word.HighlightColor.yellow,
            "green": Word.HighlightColor.green,
            "cyan": Word.HighlightColor.turquoise,
            "turquoise": Word.HighlightColor.turquoise,
            "magenta": Word.HighlightColor.pink,
            "pink": Word.HighlightColor.pink,
            "blue": Word.HighlightColor.blue,
            "red": Word.HighlightColor.red,
            "darkblue": Word.HighlightColor.darkBlue,
            "darkcyan": Word.HighlightColor.darkCyan,
            "darkgreen": Word.HighlightColor.darkGreen,
            "darkmagenta": Word.HighlightColor.darkMagenta,
            "darkred": Word.HighlightColor.darkRed,
            "darkyellow": Word.HighlightColor.darkYellow,
            "gray25": Word.HighlightColor.lightGray,
            "lightgray": Word.HighlightColor.lightGray,
            "gray50": Word.HighlightColor.darkGray,
            "darkgray": Word.HighlightColor.darkGray,
            "black": Word.HighlightColor.black,
            "nohighlight": Word.HighlightColor.noHighlight,
            "none": Word.HighlightColor.noHighlight,
          };
          const color = colorMap[highlightColor.toLowerCase()];
          if (color !== undefined) {
            font.highlightColor = color;
          }
        }

        await context.sync();

        // Build confirmation message
        const applied: string[] = [];
        if (bold !== undefined) applied.push(bold ? "bold" : "not bold");
        if (italic !== undefined) applied.push(italic ? "italic" : "not italic");
        if (underline !== undefined) applied.push(underline ? "underlined" : "not underlined");
        if (strikethrough !== undefined) applied.push(strikethrough ? "strikethrough" : "no strikethrough");
        if (fontSize !== undefined) applied.push(`${fontSize}pt`);
        if (fontName !== undefined) applied.push(fontName);
        if (fontColor !== undefined) applied.push(`color #${fontColor.replace("#", "")}`);
        if (highlightColor !== undefined) applied.push(`${highlightColor} highlight`);

        return `Applied formatting: ${applied.join(", ")}.`;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
