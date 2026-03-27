import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("add_slide_animation", "Add a PowerPoint slide animation through an Open XML slide round-trip. Supports motion paths, scale emphasis, and rotation with timing control. This replaces the slide in the deck and may change slide identity.", {
  slideIndex: tool.schema.number().describe("0-based slide index."),
  shapeId: tool.schema.string().optional().describe("Preferred shape id target."),
  shapeIndex: tool.schema.number().optional().describe("0-based shape index target if shapeId is unavailable."),
  type: tool.schema.enum(["motionPath", "scale", "rotate"]).describe("Animation type."),
  start: tool.schema.enum(["onClick", "withPrevious", "afterPrevious"]).describe("When the animation starts relative to the sequence."),
  durationMs: tool.schema.number().optional().describe("Optional animation duration in milliseconds."),
  delayMs: tool.schema.number().optional().describe("Optional start delay in milliseconds."),
  repeatCount: tool.schema.number().optional().describe("Optional repeat count."),
  path: tool.schema.string().optional().describe("Motion path string such as 'M 0 0 L 0.25 0 E'. Required for motionPath."),
  pathOrigin: tool.schema.enum(["parent", "layout"]).optional().describe("Optional motion-path origin."),
  pathEditMode: tool.schema.enum(["relative", "fixed"]).optional().describe("Optional motion-path edit mode."),
  scaleXPercent: tool.schema.number().optional().describe("Relative X scale change as a percentage."),
  scaleYPercent: tool.schema.number().optional().describe("Relative Y scale change as a percentage."),
  angleDegrees: tool.schema.number().optional().describe("Rotation amount in degrees for rotate animations."),
})
