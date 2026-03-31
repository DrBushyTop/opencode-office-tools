import { invalidSlideIndexMessage, normalizeHexColor, toolFailure } from "./powerpointShared";
import { z } from "zod";

const shapeBoundsSchema = z.object({
  left: z.number(),
  top: z.number(),
  width: z.number(),
  height: z.number(),
  name: z.string(),
  id: z.string(),
});

const imageRectangleOptionsSchema = z.object({
  left: z.number(),
  top: z.number(),
  width: z.number(),
  height: z.number(),
  name: z.string().optional(),
  imageBase64: z.string(),
});

const tableValuesSchema = z.array(z.array(z.union([z.boolean(), z.number(), z.string()])));

export type ShapeBounds = z.infer<typeof shapeBoundsSchema>;
export type ImageRectangleOptions = z.infer<typeof imageRectangleOptionsSchema>;

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

export async function getSlideById(context: PowerPoint.RequestContext, slideId: string) {
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  const slideIndex = slides.items.findIndex((slide) => slide.id === slideId);
  if (slideIndex < 0) {
    throw new Error(`Slide ${JSON.stringify(slideId)} was not found in the current presentation.`);
  }
  const slide = slides.items[slideIndex];
  return { slide, slideIndex };
}

export async function fetchImageUrlAsBase64(imageUrl: string) {
  let parsedUrl: URL;
  try {
    parsedUrl = new URL(imageUrl);
  } catch {
    throw new Error("imageUrl must be a valid HTTPS URL.");
  }

  if (parsedUrl.protocol !== "https:") {
    throw new Error("imageUrl must use HTTPS.");
  }

  const response = await fetch(imageUrl);
  if (!response.ok) {
    throw new Error(`Failed to fetch image: ${response.status} ${response.statusText}`);
  }

  const contentType = response.headers.get("content-type") || "";
  if (!contentType.toLowerCase().startsWith("image/")) {
    throw new Error("imageUrl did not return an image content type.");
  }

  const contentLengthHeader = response.headers.get("content-length");
  const contentLength = contentLengthHeader ? Number(contentLengthHeader) : 0;
  if (Number.isFinite(contentLength) && contentLength > 10 * 1024 * 1024) {
    throw new Error("imageUrl is too large.");
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
  return shapeBoundsSchema.parse({
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
    name: shape.name,
    id: shape.id,
  });
}

export function toPowerPointTableValues(values?: Array<Array<boolean | number | string>>) {
  return tableValuesSchema.parse(values || []).map((row) => row.map((cell) => String(cell ?? "")));
}

export function createImageRectangle(
  slide: PowerPoint.Slide,
  options: ImageRectangleOptions,
) {
  const parsedOptions = imageRectangleOptionsSchema.parse(options);
  const shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
    left: parsedOptions.left,
    top: parsedOptions.top,
    width: parsedOptions.width,
    height: parsedOptions.height,
  });
  shape.fill.setImage(parsedOptions.imageBase64);
  shape.lineFormat.visible = false;
  if (parsedOptions.name) shape.name = parsedOptions.name;
  return shape;
}

export function defaultChartPalette(index: number) {
  const palette = ["#1d3557", "#457b9d", "#2a9d8f", "#e9c46a", "#f4a261", "#e76f51"];
  return normalizeHexColor(palette[index % palette.length]);
}

export function failureResult(error: unknown) {
  return toolFailure(error instanceof Error ? error.message : String(error));
}
