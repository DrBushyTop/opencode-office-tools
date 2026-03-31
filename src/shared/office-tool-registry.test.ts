import { describe, expect, it } from "vitest";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import rawRegistry from "./office-tool-registry.json";
import { officeToolRegistrySourceSchema } from "./office-metadata";
import { getOfficeToolNames, isReadOnlyOfficeTool, officeToolDefinitions, officeToolRegistry } from "./office-tool-registry";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

describe("office tool registry", () => {
  it("derives definitions and permissions from one registry", () => {
    expect(() => officeToolRegistrySourceSchema.parse(rawRegistry)).not.toThrow();
    expect(officeToolDefinitions.get_document_part.description).toBe(officeToolRegistry.get_document_part.description);
    expect(isReadOnlyOfficeTool("get_document_part")).toBe(true);
    expect(isReadOnlyOfficeTool("set_document_part")).toBe(false);
    expect(isReadOnlyOfficeTool("find_document_text")).toBe(true);
    expect(isReadOnlyOfficeTool("set_document_range")).toBe(false);
  });

  it("keeps generated .opencode wrappers aligned with the registry", () => {
    const toolsDir = path.resolve(__dirname, "../../.opencode/tools");
    const wrapperFiles = fs.readdirSync(toolsDir).filter((name) => name.endsWith(".ts")).sort();
    const registryFiles = Object.keys(officeToolRegistry).map((name) => `${name}.ts`).sort();
    const getDocumentPartWrapper = fs.readFileSync(path.join(toolsDir, "get_document_part.ts"), "utf8");
    const manageRangeWrapper = fs.readFileSync(path.join(toolsDir, "manage_range.ts"), "utf8");
    const manageSlideWrapper = fs.readFileSync(path.join(toolsDir, "manage_slide.ts"), "utf8");
    const manageSlideShapesWrapper = fs.readFileSync(path.join(toolsDir, "manage_slide_shapes.ts"), "utf8");
    const createSlideFromTemplateWrapper = fs.readFileSync(path.join(toolsDir, "create_slide_from_template.ts"), "utf8");
    const editSlideWithCodeWrapper = fs.readFileSync(path.join(toolsDir, "edit_slide_with_code.ts"), "utf8");
    const manageSlideChartWrapper = fs.readFileSync(path.join(toolsDir, "manage_slide_chart.ts"), "utf8");
    const getNotebookOverviewWrapper = fs.readFileSync(path.join(toolsDir, "get_notebook_overview.ts"), "utf8");

    expect(wrapperFiles).toEqual(registryFiles);
    expect(getDocumentPartWrapper).toContain('export default word("get_document_part"');
    expect(getDocumentPartWrapper).toContain('tool.schema.enum(["summary", "text", "html"])');
    expect(wrapperFiles).toContain("get_document_range.ts");
    expect(wrapperFiles).toContain("set_document_range.ts");
    expect(wrapperFiles).toContain("find_document_text.ts");
    expect(wrapperFiles).toContain("get_document_targets.ts");
    expect(manageRangeWrapper).toContain('export default excel("manage_range"');
    expect(manageRangeWrapper).toContain('tool.schema.enum(["clear", "insert", "delete", "copy", "fill", "sort", "filter"])');
    expect(manageSlideWrapper).toContain('export default powerpoint("manage_slide"');
    expect(manageSlideWrapper).toContain('tool.schema.enum(["create", "duplicate", "delete", "move", "clear"])');
    expect(editSlideWithCodeWrapper).toContain('export default powerpoint("edit_slide_with_code"');
    expect(manageSlideShapesWrapper).toContain('export default powerpoint("manage_slide_shapes"');
    expect(manageSlideShapesWrapper).toContain('tool.schema.enum(["create", "update", "delete", "group", "ungroup"])');
    expect(manageSlideShapesWrapper).toContain('tool.schema.enum(["textBox", "geometricShape", "line"])');
    expect(manageSlideShapesWrapper).toContain('tool.schema.enum(["Straight", "Elbow", "Curve"])');
    expect(manageSlideChartWrapper).toContain('tool.schema.enum(["create", "update", "delete"])');
    expect(manageSlideChartWrapper).toContain('tool.schema.object({');
    expect(createSlideFromTemplateWrapper).toContain('export default powerpoint("create_slide_from_template"');
    expect(getNotebookOverviewWrapper).toContain('export default onenote("get_notebook_overview"');
    expect(getOfficeToolNames("word")).toContain("get_document_part");
    expect(getOfficeToolNames("word")).toContain("get_document_range");
    expect(getOfficeToolNames("powerpoint")).toContain("manage_slide");
    expect(getOfficeToolNames("powerpoint")).toContain("edit_slide_with_code");
    expect(getOfficeToolNames("powerpoint")).toContain("manage_slide_shapes");
    expect(getOfficeToolNames("powerpoint")).toContain("manage_slide_media");
    expect(getOfficeToolNames("powerpoint")).toContain("create_slide_from_template");
    expect(getOfficeToolNames("onenote")).toContain("get_notebook_overview");
  });
});
