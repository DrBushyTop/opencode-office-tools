import type { Tool } from "@github/copilot-sdk";
import TurndownService from "turndown";

const turndown = new TurndownService({
  headingStyle: "atx",
  codeBlockStyle: "fenced",
});

// Remove script, style, nav, footer, and other non-content elements
turndown.remove(["script", "style", "nav", "footer", "header", "aside", "noscript", "iframe"]);

export const webFetch: Tool = {
  name: "fetch_web_page",
  description: "Fetch content from a URL (GET only). Returns the page content as markdown. Useful for getting web page content, API data, etc.",
  parameters: {
    type: "object",
    properties: {
      url: {
        type: "string",
        description: "The URL to fetch.",
      },
    },
    required: ["url"],
  },
  handler: async (args) => {
    const { url } = args as { url: string };
    
    try {
      // Use server proxy to avoid CORS
      const response = await fetch(`/api/fetch?url=${encodeURIComponent(url)}`);
      
      if (!response.ok) {
        return `HTTP ${response.status}: ${response.statusText}`;
      }
      
      const html = await response.text();
      
      // Check if it's HTML content
      const contentType = response.headers.get("content-type") || "";
      if (!contentType.includes("text/html") && !html.trim().startsWith("<")) {
        // Return non-HTML content as-is (e.g., JSON, plain text)
        if (html.length > 50000) {
          return html.slice(0, 50000) + "\n\n[Truncated - response exceeded 50KB]";
        }
        return html;
      }

      // Extract body content if present
      const bodyMatch = html.match(/<body[^>]*>([\s\S]*)<\/body>/i);
      const bodyHtml = bodyMatch ? bodyMatch[1] : html;

      // Convert to markdown
      const markdown = turndown.turndown(bodyHtml);
      
      // Clean up excessive whitespace
      const cleaned = markdown
        .replace(/\n{3,}/g, "\n\n")
        .trim();

      if (cleaned.length > 50000) {
        return cleaned.slice(0, 50000) + "\n\n[Truncated - response exceeded 50KB]";
      }
      return cleaned;
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
