import { describe, expect, it } from "vitest";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { getOfficeToolNames, isReadOnlyOfficeTool, officeToolDefinitions, officeToolRegistry } from "./office-tool-registry";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

describe("office tool registry", () => {
  it("derives definitions and permissions from one registry", () => {
    expect(officeToolDefinitions.get_document_part.description).toBe(officeToolRegistry.get_document_part.description);
    expect(isReadOnlyOfficeTool("get_document_part")).toBe(true);
    expect(isReadOnlyOfficeTool("set_document_part")).toBe(false);
  });

  it("keeps generated .opencode wrappers aligned with the registry", () => {
    const toolsDir = path.resolve(__dirname, "../../.opencode/tools");
    const wrapperFiles = fs.readdirSync(toolsDir).filter((name) => name.endsWith(".ts")).sort();
    const registryFiles = Object.keys(officeToolRegistry).map((name) => `${name}.ts`).sort();
    const getDocumentPartWrapper = fs.readFileSync(path.join(toolsDir, "get_document_part.ts"), "utf8");

    expect(wrapperFiles).toEqual(registryFiles);
    expect(getDocumentPartWrapper).toContain('export default word("get_document_part"');
    expect(getDocumentPartWrapper).toContain('tool.schema.enum(["summary", "text", "html"])');
    expect(getOfficeToolNames("word")).toContain("get_document_part");
  });
});
