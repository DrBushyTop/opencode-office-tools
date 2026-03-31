import { describe, expect, it } from "vitest"
import { __testFormat } from "./office-tool"

describe("office-tool formatting", () => {
  it("returns structured JSON when a tool result has fields beyond textResultForLlm", () => {
    const formatted = __testFormat("list_slide_shapes", {
      textResultForLlm: "Listed 2 shapes on slide 1.",
      resultType: "success",
      slideIndex: 0,
      shapes: [{ ref: "slide-id:1/shape:10" }],
      toolTelemetry: { noisy: true },
    })

    expect(formatted).toContain('"slideIndex": 0')
    expect(formatted).toContain('"shapes"')
    expect(formatted).not.toContain('"toolTelemetry"')
  })

  it("returns plain text when the result only contains textResultForLlm", () => {
    const formatted = __testFormat("get_slide_image", {
      textResultForLlm: "Captured image of slide 1.",
    })

    expect(formatted).toBe("Captured image of slide 1.")
  })
})
