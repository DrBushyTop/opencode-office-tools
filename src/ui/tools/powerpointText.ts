export async function loadTextFrames(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<PowerPoint.TextFrame[]> {
  const frames = shapes.map((shape) => shape.getTextFrameOrNullObject());

  for (const frame of frames) {
    frame.load(["isNullObject", "hasText"]);
  }
  await context.sync();

  for (const frame of frames) {
    if (!frame.isNullObject && frame.hasText) {
      frame.textRange.load("text");
    }
  }
  await context.sync();

  return frames;
}

export async function loadShapeTexts(
  context: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<string[]> {
  const frames = await loadTextFrames(context, shapes);

  return frames.map((frame): string => {
    if (frame.isNullObject || !frame.hasText) return "";
    return frame.textRange.text || "";
  });
}
