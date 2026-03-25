export type OfficeToolHost = "word" | "powerpoint" | "excel";

export interface OfficeToolDefinition {
  name: string;
  description: string;
  parameters: Record<string, unknown>;
  hosts: OfficeToolHost[];
}

export const officeToolDefinitions: Record<string, OfficeToolDefinition> = {
  get_document_overview: {
    name: "get_document_overview",
    description: "Get a structural overview of the active Word document.",
    parameters: { type: "object", properties: {} },
    hosts: ["word"],
  },
  get_document_content: {
    name: "get_document_content",
    description: "Read the current Word document.",
    parameters: { type: "object", properties: {} },
    hosts: ["word"],
  },
  get_document_section: {
    name: "get_document_section",
    description: "Read a specific Word document section by heading.",
    parameters: {
      type: "object",
      properties: {
        headingText: { type: "string", description: "Heading text to search for." },
        includeSubsections: { type: "boolean", description: "Include nested subsections." },
      },
      required: ["headingText"],
    },
    hosts: ["word"],
  },
  set_document_content: {
    name: "set_document_content",
    description: "Replace the current Word document with new HTML content.",
    parameters: {
      type: "object",
      properties: {
        html: { type: "string", description: "HTML to write into the document." },
      },
      required: ["html"],
    },
    hosts: ["word"],
  },
  get_selection: {
    name: "get_selection",
    description: "Read the current Word selection as OOXML.",
    parameters: { type: "object", properties: {} },
    hosts: ["word"],
  },
  get_selection_text: {
    name: "get_selection_text",
    description: "Read the current Word selection as plain text.",
    parameters: { type: "object", properties: {} },
    hosts: ["word"],
  },
  insert_content_at_selection: {
    name: "insert_content_at_selection",
    description: "Insert HTML content at the current Word selection.",
    parameters: {
      type: "object",
      properties: {
        html: { type: "string", description: "HTML to insert." },
        location: {
          type: "string",
          enum: ["replace", "before", "after", "start", "end"],
          description: "Where to insert relative to the current selection.",
        },
      },
      required: ["html"],
    },
    hosts: ["word"],
  },
  find_and_replace: {
    name: "find_and_replace",
    description: "Find and replace text throughout the active Word document.",
    parameters: {
      type: "object",
      properties: {
        find: { type: "string", description: "Text to find." },
        replace: { type: "string", description: "Replacement text." },
        matchCase: { type: "boolean", description: "Match case exactly." },
        matchWholeWord: { type: "boolean", description: "Only match whole words." },
      },
      required: ["find", "replace"],
    },
    hosts: ["word"],
  },
  insert_table: {
    name: "insert_table",
    description: "Insert a table at the current Word selection.",
    parameters: {
      type: "object",
      properties: {
        data: {
          type: "array",
          items: { type: "array", items: { type: "string" } },
          description: "Two-dimensional array of table cell values.",
        },
        hasHeader: { type: "boolean", description: "Treat the first row as a header row." },
        style: { type: "string", enum: ["grid", "striped", "plain"], description: "Table style." },
      },
      required: ["data"],
    },
    hosts: ["word"],
  },
  apply_style_to_selection: {
    name: "apply_style_to_selection",
    description: "Apply formatting styles to the current Word selection.",
    parameters: {
      type: "object",
      properties: {
        bold: { type: "boolean" },
        italic: { type: "boolean" },
        underline: { type: "boolean" },
        strikethrough: { type: "boolean" },
        fontSize: { type: "number" },
        fontName: { type: "string" },
        fontColor: { type: "string" },
        highlightColor: { type: "string" },
      },
    },
    hosts: ["word"],
  },
};
