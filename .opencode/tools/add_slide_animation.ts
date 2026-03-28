import { tool } from "@opencode-ai/plugin"
import { powerpoint } from "../lib/office-powerpoint"

export default powerpoint("add_slide_animation", "Add a PowerPoint slide animation through an Open XML slide round-trip. Supports motion paths, scale emphasis, and rotation with timing control. Also supports entrance animations: appear (instant), fade (opacity), flyIn (from direction), wipe (reveal), zoomIn (scale in), floatIn (float up with fade), riseUp (rise from bottom), peekIn (fade with slight upward slide), and growAndTurn (fade with bounce from below). Entrance animations make shapes start hidden and reveal them. Emphasis color animations (complementaryColor, changeFillColor, changeLineColor) smoothly transition a shape's fill or line color. Use afterPrevious with delayMs for staggered reveal sequences. This replaces the slide in the deck and may change slide identity.", {
  slideIndex: tool.schema.number().describe("0-based slide index."),
  shapeId: tool.schema.union([tool.schema.string(), tool.schema.array(tool.schema.union([tool.schema.string(), tool.schema.number()]))]).optional().describe("Shape id target(s). Pass a single id or an array of ids to animate multiple shapes in one call. When an array is provided, the first shape uses the specified start trigger and the rest use withPrevious."),
  shapeIndex: tool.schema.number().optional().describe("0-based shape index target if shapeId is unavailable. Only for single-shape animations."),
  type: tool.schema.enum(["motionPath", "scale", "rotate", "appear", "fade", "flyIn", "wipe", "zoomIn", "floatIn", "riseUp", "peekIn", "growAndTurn", "complementaryColor", "changeFillColor", "changeLineColor"]).describe("Animation type. Entrance types (appear, fade, flyIn, wipe, zoomIn, floatIn, riseUp, peekIn, growAndTurn) start shapes hidden and reveal them. Emphasis color types (complementaryColor, changeFillColor, changeLineColor) animate a shape's color."),
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
  toColor: tool.schema.string().optional().describe("Target color for emphasis color animations. Hex without # (e.g. 'FF0000') or theme scheme name (e.g. 'accent2', 'dk1'). Required for complementaryColor, changeFillColor, changeLineColor."),
  colorSpace: tool.schema.enum(["hsl", "rgb"]).optional().describe("Color interpolation space for emphasis color animations. 'hsl' (default) gives smoother transitions."),
})
