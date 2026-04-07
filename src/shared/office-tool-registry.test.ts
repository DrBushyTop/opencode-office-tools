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
    const addSlideAnimationWrapper = fs.readFileSync(path.join(toolsDir, "add_slide_animation.ts"), "utf8");
    const executeOfficeJsWrapper = fs.readFileSync(path.join(toolsDir, "execute_office_js.ts"), "utf8");
    const listSlideShapesWrapper = fs.readFileSync(path.join(toolsDir, "list_slide_shapes.ts"), "utf8");
    const editSlideChartWrapper = fs.readFileSync(path.join(toolsDir, "edit_slide_chart.ts"), "utf8");
    const createSlideFromLayoutWrapper = fs.readFileSync(path.join(toolsDir, "create_slide_from_layout.ts"), "utf8");
    const getDocumentPartWrapper = fs.readFileSync(path.join(toolsDir, "get_document_part.ts"), "utf8");
    const manageRangeWrapper = fs.readFileSync(path.join(toolsDir, "manage_range.ts"), "utf8");
    const manageSlideWrapper = fs.readFileSync(path.join(toolsDir, "manage_slide.ts"), "utf8");
    const manageSlideShapesWrapper = fs.readFileSync(path.join(toolsDir, "manage_slide_shapes.ts"), "utf8");

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
    expect(manageSlideWrapper).toContain('tool.schema.enum(["create", "delete", "move", "clear"])');
    expect(addSlideAnimationWrapper).toContain('tool.schema.union([tool.schema.string(), tool.schema.number(), tool.schema.array(tool.schema.union([tool.schema.string(), tool.schema.number()]))])');
    expect(wrapperFiles).toContain("execute_office_js.ts");
    expect(wrapperFiles).toContain("list_slide_shapes.ts");
    expect(wrapperFiles).toContain("read_slide_text.ts");
    expect(wrapperFiles).toContain("edit_slide_text.ts");
    expect(wrapperFiles).toContain("edit_slide_xml.ts");
    expect(wrapperFiles).toContain("edit_slide_chart.ts");
    expect(wrapperFiles).toContain("edit_slide_master.ts");
    expect(wrapperFiles).toContain("get_slide_layout_details.ts");
    expect(wrapperFiles).toContain("list_slide_layouts.ts");
    expect(wrapperFiles).toContain("duplicate_slide.ts");
    expect(wrapperFiles).toContain("create_slide_from_layout.ts");
    expect(wrapperFiles).not.toContain("add_slide_from_code.ts");
    expect(executeOfficeJsWrapper).toContain('export default powerpoint("execute_office_js"');
    expect(listSlideShapesWrapper).toContain('export default powerpoint("list_slide_shapes"');
    expect(editSlideChartWrapper).toContain('export default powerpoint("edit_slide_chart"');
    expect(editSlideChartWrapper).toContain('tool.schema.enum(["create", "update", "delete"])');
    expect(createSlideFromLayoutWrapper).toContain('export default powerpoint("create_slide_from_layout"');
    expect(manageSlideShapesWrapper).toContain('export default powerpoint("manage_slide_shapes"');
    expect(manageSlideShapesWrapper).toContain('tool.schema.enum(["create", "update", "delete", "group", "ungroup"])');
    expect(manageSlideShapesWrapper).toContain('tool.schema.enum(["textBox", "geometricShape", "line"])');
    expect(manageSlideShapesWrapper).toContain('tool.schema.enum(["Straight", "Elbow", "Curve"])');
    expect(wrapperFiles).not.toContain("get_slide_shapes.ts");
    expect(wrapperFiles).not.toContain("manage_slide_chart.ts");
    expect(wrapperFiles).not.toContain("insert_business_layout.ts");
    expect(wrapperFiles).not.toContain("create_slide_from_template.ts");
    expect(getOfficeToolNames("word")).toContain("get_document_part");
    expect(getOfficeToolNames("word")).toContain("get_document_range");
    expect(getOfficeToolNames("powerpoint")).toContain("add_slide_animation");
    expect(getOfficeToolNames("powerpoint")).toContain("clear_slide_animations");
    expect(getOfficeToolNames("powerpoint")).toContain("get_slide_animations");
    expect(getOfficeToolNames("powerpoint")).toContain("execute_office_js");
    expect(getOfficeToolNames("powerpoint")).toContain("get_slide_notes");
    expect(getOfficeToolNames("powerpoint")).toContain("get_slide_transition");
    expect(getOfficeToolNames("powerpoint")).toContain("list_slide_shapes");
    expect(getOfficeToolNames("powerpoint")).toContain("read_slide_text");
    expect(getOfficeToolNames("powerpoint")).toContain("edit_slide_text");
    expect(getOfficeToolNames("powerpoint")).toContain("edit_slide_xml");
    expect(getOfficeToolNames("powerpoint")).toContain("edit_slide_chart");
    expect(getOfficeToolNames("powerpoint")).toContain("edit_slide_master");
    expect(getOfficeToolNames("powerpoint")).toContain("get_slide_layout_details");
    expect(getOfficeToolNames("powerpoint")).toContain("list_slide_layouts");
    expect(getOfficeToolNames("powerpoint")).toContain("duplicate_slide");
    expect(getOfficeToolNames("powerpoint")).toContain("create_slide_from_layout");
    expect(getOfficeToolNames("powerpoint")).toContain("manage_slide");
    expect(getOfficeToolNames("powerpoint")).toContain("manage_slide_shapes");
    expect(getOfficeToolNames("powerpoint")).toContain("manage_slide_media");
    expect(getOfficeToolNames("powerpoint")).toContain("manage_slide_table");
    expect(getOfficeToolNames("powerpoint")).toContain("set_slide_notes");
    expect(getOfficeToolNames("powerpoint")).toContain("set_slide_transition");
    expect(getOfficeToolNames("powerpoint")).not.toContain("add_slide_from_code");
    expect(getOfficeToolNames("powerpoint")).not.toContain("get_slide_shapes");
    expect(getOfficeToolNames("powerpoint")).not.toContain("manage_slide_chart");
    expect(getOfficeToolNames("powerpoint")).not.toContain("insert_business_layout");
    expect(getOfficeToolNames("powerpoint")).not.toContain("create_slide_from_template");
  });

  it("keeps active prompt assets free of removed PowerPoint tool names", () => {
    const removedToolNamePattern = /get_slide_shapes|manage_slide_chart|insert_business_layout|create_slide_from_template|add_slide_from_code/;
    const activePromptAssets = [
      path.resolve(__dirname, "../../.opencode/agents/powerpoint.md"),
      path.resolve(__dirname, "../../.opencode/agents/visual-qa.md"),
      path.resolve(__dirname, "../../TOOLS_CATALOG.md"),
    ];

    for (const assetPath of activePromptAssets) {
      expect(fs.readFileSync(assetPath, "utf8")).not.toMatch(removedToolNamePattern);
    }
  });
});
