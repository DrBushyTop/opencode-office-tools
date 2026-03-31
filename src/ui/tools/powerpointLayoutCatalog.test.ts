import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from "@xmldom/xmldom";
import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import {
  lookupPresentationLayoutMetadata,
  parsePresentationLayoutCatalogFromBase64,
  resolveSlideLayoutMetadata,
} from "./powerpointLayoutCatalog";

function createPresentationBase64(entries: Record<string, string>) {
  let binary = "";
  zipSync(Object.fromEntries(
    Object.entries(entries).map(([path, contents]) => [path, strToU8(contents)]),
  )).forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary);
}

if (typeof DOMParser === "undefined") {
  Object.assign(globalThis, {
    DOMParser: XmldomParser,
    XMLSerializer: XmldomSerializer,
  });
}

describe("powerpointLayoutCatalog", () => {
  it("parses layout names and types from presentation XML", () => {
    const base64 = createPresentationBase64({
      "ppt/presentation.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:sldMasterIdLst>
            <p:sldMasterId id="2147483698" r:id="rId1"/>
          </p:sldMasterIdLst>
        </p:presentation>`,
      "ppt/_rels/presentation.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
        </Relationships>`,
      "ppt/slideMasters/slideMaster1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <p:cSld name="Zure Theme Light"/>
          <p:sldLayoutIdLst>
            <p:sldLayoutId id="2147483700" r:id="rId1"/>
            <p:sldLayoutId id="2147483701" r:id="rId2"/>
          </p:sldLayoutIdLst>
        </p:sldMaster>`,
      "ppt/slideMasters/_rels/slideMaster1.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
        </Relationships>`,
      "ppt/slideLayouts/slideLayout1.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" matchingName="Light: Title and Subtitle" type="title">
          <p:cSld name="Ignored because matchingName wins"/>
        </p:sldLayout>`,
      "ppt/slideLayouts/slideLayout2.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="objTx">
          <p:cSld name="Content With Caption"/>
        </p:sldLayout>`,
    });

    const catalog = parsePresentationLayoutCatalogFromBase64(base64);

    expect(catalog).toEqual({
      slideMasters: [
        {
          openXmlId: "2147483698",
          slideMasterName: "Zure Theme Light",
          layouts: [
            { openXmlId: "2147483700", layoutName: "Light: Title and Subtitle", layoutType: "Title" },
            { openXmlId: "2147483701", layoutName: "Content With Caption", layoutType: "ContentWithCaption" },
          ],
        },
      ],
    });
  });

  it("matches Office ids by the Open XML id prefix and resolves fallback metadata", () => {
    const catalog = {
      slideMasters: [
        {
          openXmlId: "2147483698",
          slideMasterName: "Zure Theme Light",
          layouts: [
            { openXmlId: "2147483700", layoutName: "Light: Title and Subtitle", layoutType: "Title" },
          ],
        },
      ],
    };

    expect(lookupPresentationLayoutMetadata(catalog, {
      slideMasterId: "2147483698#626277182",
      layoutId: "2147483700#2384306777",
    })).toEqual({
      slideMasterName: "Zure Theme Light",
      layoutName: "Light: Title and Subtitle",
      layoutType: "Title",
    });

    expect(resolveSlideLayoutMetadata("", "", {
      layoutName: "Light: Title and Subtitle",
      layoutType: "Title",
    })).toEqual({
      layoutName: "Light: Title and Subtitle",
      layoutType: "Title",
    });
  });
});
