import { getDocumentContent } from "./getDocumentContent";
import { setDocumentContent } from "./setDocumentContent";
import { getSelection } from "./getSelection";
import { getPresentationContent } from "./getPresentationContent";
import { getPresentationOverview } from "./getPresentationOverview";
import { getSlideImage } from "./getSlideImage";
import { setPresentationContent } from "./setPresentationContent";
import { addSlideFromCode } from "./addSlideFromCode";
import { clearSlide } from "./clearSlide";
import { updateSlideShape } from "./updateSlideShape";
import { getWorkbookContent } from "./getWorkbookContent";
import { setWorkbookContent } from "./setWorkbookContent";
import { getSelectedRange } from "./getSelectedRange";
import { setSelectedRange } from "./setSelectedRange";
import { getWorkbookInfo } from "./getWorkbookInfo";

// New Word tools
import { getDocumentOverview } from "./getDocumentOverview";
import { getSelectionText } from "./getSelectionText";
import { insertContentAtSelection } from "./insertContentAtSelection";
import { findAndReplace } from "./findAndReplace";
import { getDocumentSection } from "./getDocumentSection";
import { insertTable } from "./insertTable";
import { applyStyleToSelection } from "./applyStyleToSelection";
import { getDocumentPart } from "./getDocumentPart";
import { setDocumentPart } from "./setDocumentPart";

// New PowerPoint tools
import { addSlideAnimation } from "./addSlideAnimation";
import { clearSlideAnimations } from "./clearSlideAnimations";
import { getSlideNotes } from "./getSlideNotes";
import { getSlideTransition } from "./getSlideTransition";
import { setSlideNotes } from "./setSlideNotes";
import { setSlideTransition } from "./setSlideTransition";
import { duplicateSlide } from "./duplicateSlide";
import { getPresentationStructure } from "./getPresentationStructure";
import { getSlideShapes } from "./getSlideShapes";
import { setSlideShapeProperties } from "./setSlideShapeProperties";
import { deleteSlide } from "./deleteSlide";
import { moveSlide } from "./moveSlide";

// New Excel tools
import { getWorkbookOverview } from "./getWorkbookOverview";
import { findAndReplaceCells } from "./findAndReplaceCells";
import { insertChart } from "./insertChart";
import { applyCellFormatting } from "./applyCellFormatting";
import { createNamedRange } from "./createNamedRange";
import { manageWorksheet } from "./manageWorksheet";
import { manageTable } from "./manageTable";
import { getOfficeToolNames } from "../../shared/office-tool-definitions";

const officeToolHandlers = {
  [getDocumentOverview.name]: getDocumentOverview,
  [getDocumentContent.name]: getDocumentContent,
  [getDocumentPart.name]: getDocumentPart,
  [getDocumentSection.name]: getDocumentSection,
  [setDocumentContent.name]: setDocumentContent,
  [setDocumentPart.name]: setDocumentPart,
  [getSelection.name]: getSelection,
  [getSelectionText.name]: getSelectionText,
  [insertContentAtSelection.name]: insertContentAtSelection,
  [findAndReplace.name]: findAndReplace,
  [insertTable.name]: insertTable,
  [applyStyleToSelection.name]: applyStyleToSelection,
  [getWorkbookOverview.name]: getWorkbookOverview,
  [getWorkbookInfo.name]: getWorkbookInfo,
  [getWorkbookContent.name]: getWorkbookContent,
  [setWorkbookContent.name]: setWorkbookContent,
  [getSelectedRange.name]: getSelectedRange,
  [setSelectedRange.name]: setSelectedRange,
  [findAndReplaceCells.name]: findAndReplaceCells,
  [insertChart.name]: insertChart,
  [applyCellFormatting.name]: applyCellFormatting,
  [createNamedRange.name]: createNamedRange,
  [manageWorksheet.name]: manageWorksheet,
  [manageTable.name]: manageTable,
  [getPresentationOverview.name]: getPresentationOverview,
  [getPresentationStructure.name]: getPresentationStructure,
  [getPresentationContent.name]: getPresentationContent,
  [getSlideImage.name]: getSlideImage,
  [getSlideShapes.name]: getSlideShapes,
  [addSlideAnimation.name]: addSlideAnimation,
  [clearSlideAnimations.name]: clearSlideAnimations,
  [getSlideNotes.name]: getSlideNotes,
  [getSlideTransition.name]: getSlideTransition,
  [setPresentationContent.name]: setPresentationContent,
  [addSlideFromCode.name]: addSlideFromCode,
  [clearSlide.name]: clearSlide,
  [updateSlideShape.name]: updateSlideShape,
  [setSlideShapeProperties.name]: setSlideShapeProperties,
  [deleteSlide.name]: deleteSlide,
  [moveSlide.name]: moveSlide,
  [setSlideNotes.name]: setSlideNotes,
  [setSlideTransition.name]: setSlideTransition,
  [duplicateSlide.name]: duplicateSlide,
};

export const allOfficeTools = Object.values(officeToolHandlers);

export const wordTools = [
  getDocumentOverview,
  getDocumentContent,
  getDocumentPart,
  getDocumentSection,
  setDocumentContent,
  setDocumentPart,
  getSelection,
  getSelectionText,
  insertContentAtSelection,
  findAndReplace,
  insertTable,
  applyStyleToSelection,
];

export const powerpointTools = [
  getPresentationOverview,
  getPresentationStructure,
  getPresentationContent,
  getSlideImage,
  getSlideShapes,
  addSlideAnimation,
  clearSlideAnimations,
  getSlideNotes,
  getSlideTransition,
  setPresentationContent,
  addSlideFromCode,
  clearSlide,
  updateSlideShape,
  setSlideShapeProperties,
  deleteSlide,
  moveSlide,
  setSlideNotes,
  setSlideTransition,
  duplicateSlide,
];

export const excelTools = [
  getWorkbookOverview,
  getWorkbookInfo,
  getWorkbookContent,
  setWorkbookContent,
  getSelectedRange,
  setSelectedRange,
  findAndReplaceCells,
  insertChart,
  applyCellFormatting,
  createNamedRange,
  manageWorksheet,
  manageTable,
];

export function getToolsForHost(host: typeof Office.HostType[keyof typeof Office.HostType]) {
  const registryHost = host === Office.HostType.Word
    ? "word"
    : host === Office.HostType.PowerPoint
      ? "powerpoint"
      : host === Office.HostType.Excel
        ? "excel"
        : null;

  if (!registryHost) return [];

  return getOfficeToolNames(registryHost).map((name) => {
    const tool = officeToolHandlers[name as keyof typeof officeToolHandlers];
    if (!tool) {
      throw new Error(`Missing Office tool handler for '${name}'`);
    }
    return tool;
  });
}

export function getOfficeToolExecutor(host: typeof Office.HostType[keyof typeof Office.HostType]) {
  const tools = getToolsForHost(host);
  const map = new Map(tools.map((tool) => [tool.name, tool.handler]));

  return {
    async execute(toolName: string, args: Record<string, unknown>) {
      const handler = map.get(toolName);
      if (!handler) {
        throw new Error(`Tool '${toolName}' is not available for this host`);
      }

      const result = await handler(args as never, {
        sessionId: "office-bridge",
        toolCallId: crypto.randomUUID(),
        toolName,
        arguments: args,
      } as never);

      if (typeof result === "string") return result;
      if ((result as any).resultType === "failure") {
        throw new Error((result as any).error || (result as any).textResultForLlm || "Tool execution failed");
      }

      return (result as any).textResultForLlm || result;
    },
  };
}

export function getToolNamesForHost(host: "word" | "powerpoint" | "excel") {
  return getOfficeToolNames(host);
}
