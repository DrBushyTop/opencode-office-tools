import { OpenXmlPackage, parseXml, relationshipPartPath, resolveTargetPath } from "./openXmlPackage";

const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const xmlSlideLayoutTypeMap: Record<string, string> = {
  blank: "Blank",
  chart: "Chart",
  chartAndTx: "ChartAndText",
  clipArtAndTx: "ClipArtAndText",
  clipArtAndVertTx: "ClipArtAndVerticalText",
  cust: "Custom",
  dgm: "OrganizationChart",
  fourObj: "FourObjects",
  mediaAndTx: "MediaClipAndText",
  obj: "Object",
  objAndTx: "ObjectAndText",
  objAndTwoObj: "ObjectAndTwoObjects",
  objOnly: "LargeObject",
  objOverTx: "ObjectOverText",
  objTx: "ContentWithCaption",
  picTx: "PictureWithCaption",
  secHead: "SectionHeader",
  tbl: "Table",
  tx: "Text",
  txAndChart: "TextAndChart",
  txAndClipArt: "TextAndClipArt",
  txAndMedia: "TextAndMediaClip",
  txAndObj: "TextAndObject",
  txAndTwoObj: "TextAndTwoObjects",
  txOverObj: "TextOverObject",
  title: "Title",
  titleOnly: "TitleOnly",
  twoColTx: "TwoColumnText",
  twoObj: "TwoObjects",
  twoObjAndObj: "TwoObjectsAndObject",
  twoObjAndTx: "TwoObjectsAndText",
  twoObjOverTx: "TwoObjectsOverText",
  twoTxTwoObj: "TwoTextAndTwoObjects",
  vertTx: "VerticalText",
  vertTitleAndTx: "VerticalTitleAndText",
  vertTitleAndTxOverChart: "VerticalTitleAndTextOverChart",
};

export interface PresentationLayoutCatalogLayout {
  openXmlId: string;
  layoutName: string;
  layoutType: string;
}

export interface PresentationLayoutCatalogMaster {
  openXmlId: string;
  slideMasterName: string;
  layouts: PresentationLayoutCatalogLayout[];
}

export interface PresentationLayoutCatalog {
  slideMasters: PresentationLayoutCatalogMaster[];
}

function getNamespacedAttribute(element: Element, namespace: string, localName: string, qualifiedName: string) {
  return element.getAttributeNS(namespace, localName) || element.getAttribute(qualifiedName) || "";
}

function getTrimmedAttribute(element: Element | undefined | null, name: string) {
  return element?.getAttribute(name)?.trim() || "";
}

function getRelationshipTargetPath(pkg: OpenXmlPackage, sourcePartPath: string, relationshipId: string) {
  const relationshipsPath = relationshipPartPath(sourcePartPath);
  if (!pkg.has(relationshipsPath)) return "";

  const relationshipsDoc = parseXml(pkg.readText(relationshipsPath));
  const relationship = Array.from(relationshipsDoc.getElementsByTagName("Relationship"))
    .find((item) => item.getAttribute("Id") === relationshipId);
  const target = relationship?.getAttribute("Target") || "";
  return target ? resolveTargetPath(sourcePartPath, target) : "";
}

function formatSlideLayoutDisplayName(layoutType: string) {
  switch (layoutType) {
    case "TitleOnly":
      return "Title Only";
    case "TitleAndContent":
      return "Title and Content";
    case "PictureWithCaption":
      return "Picture with Caption";
    case "SectionHeader":
      return "Section Header";
    default:
      return layoutType.replace(/([a-z])([A-Z])/g, "$1 $2");
  }
}

function normalizeLayoutMetadataValue(value: string | null | undefined) {
  const trimmed = value?.trim() || "";
  return trimmed && trimmed !== "Unknown" && trimmed !== "undefined" && trimmed !== "null"
    ? trimmed
    : "";
}

function extractOfficeIdPrefix(officeId: string | null | undefined) {
  const trimmed = officeId?.trim() || "";
  if (!trimmed) return "";
  return trimmed.split("#")[0] || "";
}

function getMasterCommonSlideName(masterDoc: XMLDocument) {
  const cSld = masterDoc.getElementsByTagNameNS(NS_P, "cSld")[0];
  return getTrimmedAttribute(cSld, "name");
}

function mapXmlSlideLayoutType(xmlType: string) {
  return xmlSlideLayoutTypeMap[xmlType] || xmlType;
}

function encodeBytesAsBase64(bytes: Uint8Array) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  return btoa(binary);
}

function getFileAsync(fileType: Office.FileType, options: Office.GetFileOptions) {
  return new Promise<Office.File>((resolve, reject) => {
    Office.context.document.getFileAsync(fileType, options, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

function getSliceAsync(file: Office.File, sliceIndex: number) {
  return new Promise<Office.Slice>((resolve, reject) => {
    file.getSliceAsync(sliceIndex, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

function closeFileAsync(file: Office.File) {
  return new Promise<void>((resolve) => {
    file.closeAsync(() => resolve());
  });
}

export async function readCurrentPresentationAsBase64() {
  const file = await getFileAsync(Office.FileType.Compressed, { sliceSize: 1024 * 1024 });

  try {
    const slices: number[][] = [];
    for (let index = 0; index < file.sliceCount; index += 1) {
      const slice = await getSliceAsync(file, index);
      slices[slice.index] = Array.from(slice.data as number[]);
    }

    const totalLength = slices.reduce((total, slice) => total + slice.length, 0);
    const bytes = new Uint8Array(totalLength);
    let offset = 0;
    for (const slice of slices) {
      bytes.set(slice, offset);
      offset += slice.length;
    }
    return encodeBytesAsBase64(bytes);
  } finally {
    await closeFileAsync(file);
  }
}

export function parsePresentationLayoutCatalogFromBase64(base64: string): PresentationLayoutCatalog {
  const pkg = new OpenXmlPackage(base64);
  const presentationPath = "ppt/presentation.xml";
  const presentationDoc = parseXml(pkg.readText(presentationPath));
  const masterIdNodes = Array.from(presentationDoc.getElementsByTagNameNS(NS_P, "sldMasterId"));

  const slideMasters = masterIdNodes.map<PresentationLayoutCatalogMaster>((masterIdNode) => {
    const masterOpenXmlId = getTrimmedAttribute(masterIdNode, "id");
    const masterRelationshipId = getNamespacedAttribute(masterIdNode, NS_R, "id", "r:id");
    const masterPath = getRelationshipTargetPath(pkg, presentationPath, masterRelationshipId);
    const masterDoc = parseXml(pkg.readText(masterPath));
    const layoutIdNodes = Array.from(masterDoc.getElementsByTagNameNS(NS_P, "sldLayoutId"));

    const layouts = layoutIdNodes.map<PresentationLayoutCatalogLayout>((layoutIdNode) => {
      const layoutOpenXmlId = getTrimmedAttribute(layoutIdNode, "id");
      const layoutRelationshipId = getNamespacedAttribute(layoutIdNode, NS_R, "id", "r:id");
      const layoutPath = getRelationshipTargetPath(pkg, masterPath, layoutRelationshipId);
      const layoutDoc = parseXml(pkg.readText(layoutPath));
      const layoutRoot = layoutDoc.documentElement;
      const commonSlide = layoutDoc.getElementsByTagNameNS(NS_P, "cSld")[0];
      const layoutType = mapXmlSlideLayoutType(getTrimmedAttribute(layoutRoot, "type")) || "Unknown";
      const layoutName = getTrimmedAttribute(layoutRoot, "matchingName") || getTrimmedAttribute(commonSlide, "name") || formatSlideLayoutDisplayName(layoutType);

      return {
        openXmlId: layoutOpenXmlId,
        layoutName,
        layoutType,
      };
    });

    return {
      openXmlId: masterOpenXmlId,
      slideMasterName: getMasterCommonSlideName(masterDoc),
      layouts,
    };
  });

  return { slideMasters };
}

export async function loadPresentationLayoutCatalogFromDocument() {
  const base64 = await readCurrentPresentationAsBase64();
  return parsePresentationLayoutCatalogFromBase64(base64);
}

export function lookupPresentationLayoutMetadata(
  catalog: PresentationLayoutCatalog | null | undefined,
  options: { slideMasterId?: string; layoutId?: string; masterIndex?: number; layoutIndex?: number },
) {
  if (!catalog) return null;

  const masterById = extractOfficeIdPrefix(options.slideMasterId);
  const layoutById = extractOfficeIdPrefix(options.layoutId);
  const master = (masterById
    ? catalog.slideMasters.find((item) => item.openXmlId === masterById)
    : undefined)
    || (options.masterIndex !== undefined ? catalog.slideMasters[options.masterIndex] : undefined)
    || null;
  if (!master) return null;

  if (!extractOfficeIdPrefix(options.layoutId) && options.layoutIndex === undefined) {
    return {
      slideMasterName: master.slideMasterName,
      layoutName: "",
      layoutType: "",
    };
  }

  const layout = (layoutById
    ? master.layouts.find((item) => item.openXmlId === layoutById)
    : undefined)
    || (options.layoutIndex !== undefined ? master.layouts[options.layoutIndex] : undefined)
    || null;
  if (!layout) return null;

  return {
    slideMasterName: master.slideMasterName,
    layoutName: layout.layoutName,
    layoutType: layout.layoutType,
  };
}

export function resolveSlideLayoutMetadata(
  officeLayoutName: string | null | undefined,
  officeLayoutType: string | null | undefined,
  fallback: { layoutName?: string | null; layoutType?: string | null } | null | undefined,
) {
  const layoutType = normalizeLayoutMetadataValue(officeLayoutType)
    || normalizeLayoutMetadataValue(fallback?.layoutType)
    || "Unknown";
  const layoutName = normalizeLayoutMetadataValue(officeLayoutName)
    || normalizeLayoutMetadataValue(fallback?.layoutName)
    || (layoutType !== "Unknown" ? formatSlideLayoutDisplayName(layoutType) : "");

  return { layoutName, layoutType };
}
