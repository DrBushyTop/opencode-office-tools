import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("set_slide_transition", "Set a PowerPoint slide transition by round-tripping the slide through Open XML. This replaces the slide in the deck and may change slide identity.", {
  slideIndex: tool.schema.union([tool.schema.number(), tool.schema.array(tool.schema.number())]).optional().describe("0-based slide index or array of slide indexes. When an array is provided, the same transition is applied to each slide in sequence."),
  effect: tool.schema.enum(["none", "cut", "fade", "dissolve", "random", "randomBar", "push", "wipe", "split", "cover", "pull", "zoom"]).describe("Transition effect to apply. Use none to clear the transition."),
  speed: tool.schema.enum(["slow", "medium", "fast"]).optional().describe("Optional transition speed."),
  advanceOnClick: tool.schema.boolean().optional().describe("Whether a click advances the slide."),
  advanceAfterMs: tool.schema.number().optional().describe("Optional auto-advance delay in milliseconds."),
  durationMs: tool.schema.number().optional().describe("Optional transition duration in milliseconds for clients that support the p14 duration extension."),
  direction: tool.schema.enum(["left", "right", "up", "down", "horizontal", "vertical", "in", "out"]).optional().describe("Optional direction for effects that support it."),
  orientation: tool.schema.enum(["horizontal", "vertical"]).optional().describe("Optional orientation for split transitions."),
  throughBlack: tool.schema.boolean().optional().describe("Optional cut-through-black flag for cut transitions."),
})
