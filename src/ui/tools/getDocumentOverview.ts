import type { Tool } from "@github/copilot-sdk";

export const getDocumentOverview: Tool = {
  name: "get_document_overview",
  description: `Get a structural overview of the Word document. Use this first to understand the document before reading or editing specific sections.

Returns:
- Total word count and paragraph count
- Heading structure (H1, H2, H3 hierarchy with text)
- Table count
- List count (bulleted and numbered)
- Content control count

This is faster than reading the entire document and helps you understand what to target for edits.`,
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async () => {
    try {
      return await Word.run(async (context) => {
        const body = context.document.body;
        
        // Load basic stats
        body.load("text");
        
        // Get all paragraphs with their styles
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        
        // Get tables
        const tables = body.tables;
        tables.load("items");
        
        // Get content controls
        const contentControls = body.contentControls;
        contentControls.load("items");
        
        await context.sync();
        
        // Load paragraph details
        for (const para of paragraphs.items) {
          para.load(["text", "style", "isListItem"]);
        }
        await context.sync();
        
        // Calculate stats
        const text = body.text || "";
        const wordCount = text.trim().split(/\s+/).filter(w => w.length > 0).length;
        const paragraphCount = paragraphs.items.length;
        const tableCount = tables.items.length;
        const contentControlCount = contentControls.items.length;
        
        // Build heading structure
        const headings: string[] = [];
        let listCount = 0;
        
        for (const para of paragraphs.items) {
          const style = para.style || "";
          const paraText = (para.text || "").trim();
          
          if (para.isListItem) {
            listCount++;
          }
          
          // Check for heading styles
          if (style.match(/Heading\s*1/i) || style === "Title") {
            headings.push(`# ${paraText.substring(0, 80)}${paraText.length > 80 ? "..." : ""}`);
          } else if (style.match(/Heading\s*2/i) || style === "Subtitle") {
            headings.push(`  ## ${paraText.substring(0, 70)}${paraText.length > 70 ? "..." : ""}`);
          } else if (style.match(/Heading\s*3/i)) {
            headings.push(`    ### ${paraText.substring(0, 60)}${paraText.length > 60 ? "..." : ""}`);
          } else if (style.match(/Heading\s*[4-6]/i)) {
            headings.push(`      #### ${paraText.substring(0, 50)}${paraText.length > 50 ? "..." : ""}`);
          }
        }
        
        // Build output
        let output = `Document Overview:\n`;
        output += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
        output += `Words: ${wordCount.toLocaleString()}\n`;
        output += `Paragraphs: ${paragraphCount}\n`;
        output += `Tables: ${tableCount}\n`;
        output += `List items: ${listCount}\n`;
        if (contentControlCount > 0) {
          output += `Content controls: ${contentControlCount}\n`;
        }
        
        if (headings.length > 0) {
          output += `\nDocument Structure:\n`;
          output += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
          output += headings.join("\n");
        } else {
          output += `\n(No headings found - document may be unstructured)`;
        }
        
        return output;
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
