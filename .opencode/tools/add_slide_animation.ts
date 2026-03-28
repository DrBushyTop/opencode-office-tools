import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("add_slide_animation", "Add a PowerPoint slide animation through an Open XML slide round-trip. Supports motion paths, scale emphasis, and rotation with timing control. Also supports entrance animations: appear (instant), fade (opacity), flyIn (from direction), wipe (reveal), and zoomIn (scale in). Entrance animations make shapes start hidden and reveal them. Use afterPrevious with delayMs for staggered reveal sequences. This replaces the slide in the deck and may change slide identity.", {
  slideIndex: tool.schema.number().describe("0-based slide index."),
  shapeId: tool.schema.string().optional().describe("Preferred shape id target."),
  shapeIndex: tool.schema.number().optional().describe("0-based shape index target if shapeId is unavailable."),
  type: tool.schema.enum(["motionPath", "scale", "rotate", "appear", "fade", "flyIn", "wipe", "zoomIn"]).describe("Animation type. Entrance types (appear, fade, flyIn, wipe, zoomIn) start shapes hidden and reveal them."),
  start: tool.schema.enum(["onClick", "withPrevious", "afterPrevious"]).describe("When the animation starts relative to the sequence. Use afterPrevious with delayMs for staggered reveals."),
  durationMs: tool.schema.number().optional().describe("Optional animation duration in milliseconds. Default 1000. For appear, this is effectively instant."),
  delayMs: tool.schema.number().optional().describe("Optional start delay in milliseconds. Useful for staggered entrance sequences."),
  repeatCount: tool.schema.number().optional().describe("Optional repeat count."),
  path: tool.schema.string().optional().describe("Motion path string such as 'M 0 0 L 0.25 0 E'. Required for motionPath."),
  pathOrigin: tool.schema.enum(["parent", "layout"]).optional().describe("Optional motion-path origin."),
  pathEditMode: tool.schema.enum(["relative", "fixed"]).optional().describe("Optional motion-path edit mode."),
  scaleXPercent: tool.schema.number().optional().describe("Relative X scale change as a percentage."),
  scaleYPercent: tool.schema.number().optional().describe("Relative Y scale change as a percentage."),
  angleDegrees: tool.schema.number().optional().describe("Rotation amount in degrees for rotate animations."),
  direction: tool.schema.enum(["left", "right", "up", "down"]).optional().describe("Direction for flyIn (where the shape flies from) or wipe (reveal direction) entrance animations."),
})
