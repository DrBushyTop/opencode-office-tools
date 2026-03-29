import { invalidSlideIndexMessage, normalizeHexColor, toolFailure } from "./powerpointShared";

export async function getSlideByIndex(context: PowerPoint.RequestContext, slideIndex: number) {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  const slide = slides.items[slideIndex];
  if (!slide) {
    throw new Error(invalidSlideIndexMessage(slideIndex, slides.items.length));
  }
  return slide;
}

export async function fetchImageUrlAsBase64(imageUrl: string) {
  const response = await fetch(imageUrl);
  if (!response.ok) {
    throw new Error(`Failed to fetch image: ${response.status} ${response.statusText}`);
  }
  const blob = await response.blob();
  return await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Failed to read image response as base64."));
    reader.onload = () => {
      const value = typeof reader.result === "string" ? reader.result : "";
      resolve(value.replace(/^data:[^,]+,/, ""));
    };
    reader.readAsDataURL(blob);
  });
}

export async function getShapeBounds(shape: PowerPoint.Shape, context: PowerPoint.RequestContext) {
  shape.load(["left", "top", "width", "height", "name", "id"]);
  await context.sync();
  return {
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
    name: shape.name,
    id: shape.id,
  };
}

export function toPowerPointTableValues(values?: Array<Array<boolean | number | string>>) {
  return (values || []).map((row) => row.map((cell) => String(cell ?? "")));
}

export function createImageRectangle(
  slide: PowerPoint.Slide,
  options: { left: number; top: number; width: number; height: number; name?: string; imageBase64: string },
) {
  const shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
    left: options.left,
    top: options.top,
    width: options.width,
    height: options.height,
  });
  shape.fill.setImage(options.imageBase64);
  shape.lineFormat.visible = false;
  if (options.name) shape.name = options.name;
  return shape;
}

export function defaultChartPalette(index: number) {
  const palette = ["#1d3557", "#457b9d", "#2a9d8f", "#e9c46a", "#f4a261", "#e76f51"];
  return normalizeHexColor(palette[index % palette.length]);
}

export function failureResult(error: unknown) {
  return toolFailure(error instanceof Error ? error.message : String(error));
}
