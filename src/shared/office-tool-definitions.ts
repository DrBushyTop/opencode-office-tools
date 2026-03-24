export type OfficeToolHost = "word" | "powerpoint" | "excel";

export interface OfficeToolDefinition {
  name: string;
  description: string;
  parameters: Record<string, unknown>;
  hosts: OfficeToolHost[];
}

export const officeToolDefinitions: Record<string, OfficeToolDefinition> = {
  get_document_content: {
    name: "get_document_content",
    description: "Read the current Word document.",
    parameters: {
      type: "object",
      properties: {},
    },
    hosts: ["word"],
  },
};
