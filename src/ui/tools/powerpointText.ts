export async function loadTextFrames(context: PowerPoint.RequestContext, shapes: PowerPoint.Shape[]) {
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

export async function loadShapeTexts(context: PowerPoint.RequestContext, shapes: PowerPoint.Shape[]) {
  const frames = await loadTextFrames(context, shapes);

  return frames.map((frame) => {
    if (frame.isNullObject || !frame.hasText) return "";
    return frame.textRange.text || "";
  });
}
