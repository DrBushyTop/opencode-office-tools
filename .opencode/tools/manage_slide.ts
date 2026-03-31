import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("manage_slide", "Create, delete, move, or clear PowerPoint slides with one generic slide-management tool.", {
  action: tool.schema.enum(["create", "delete", "move", "clear"]).describe("Slide operation to perform."),
  slideIndex: tool.schema.number().optional().describe("0-based target slide index for delete, move, or clear."),
  targetIndex: tool.schema.number().optional().describe("0-based destination or insertion index for create or move."),
  slideMasterId: tool.schema.string().optional().describe("Optional slide master id for create."),
  layoutId: tool.schema.string().optional().describe("Optional layout id for create."),
  formatting: tool.schema.enum(["KeepSourceFormatting", "UseDestinationTheme"]).optional().describe("Formatting behavior hint used when the host needs to insert created content. Default KeepSourceFormatting."),
})
