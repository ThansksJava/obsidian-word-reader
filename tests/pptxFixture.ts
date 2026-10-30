import JSZip from "jszip";

const RELATIONSHIPS_NAMESPACE =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export async function createMinimalPptx(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`,
  );
  zip.file(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:rel="${OFFICE_RELATIONSHIP}"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst><p:sldId id="256" rel:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`,
  );
  zip.file(
    "ppt/_rels/presentation.xml.rels",
    relationships([
      [
        "rId1",
        `${OFFICE_RELATIONSHIP}/slide`,
        "slides/slide1.xml",
      ],
    ]),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:rel="${OFFICE_RELATIONSHIP}"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="3200" b="1"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>Hello PPTX</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="3" name="Picture" descr="Sample image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip rel:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="6096000" y="914400"/><a:ext cx="3048000" cy="2286000"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
    </p:pic>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="4" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="914400" y="3657600"/><a:ext cx="5486400" cy="1371600"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
        <a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="2743200"/><a:gridCol w="2743200"/></a:tblGrid>
          <a:tr h="685800">
            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>A1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
            <a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>B1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>
          </a:tr>
        </a:tbl>
      </a:graphicData></a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>`,
  );
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    relationships([
      [
        "rId1",
        `${OFFICE_RELATIONSHIP}/slideLayout`,
        "../slideLayouts/slideLayout1.xml",
      ],
      [
        "rId2",
        `${OFFICE_RELATIONSHIP}/image`,
        "../media/image1.png",
      ],
      [
        "rId3",
        `${OFFICE_RELATIONSHIP}/notesSlide`,
        "../notesSlides/notesSlide1.xml",
      ],
    ]),
  );
  zip.file(
    "ppt/notesSlides/notesSlide1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Discuss the quarterly trend.</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Slide Number"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld><a:t>1</a:t></a:fld></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:notes>`,
  );
  zip.file(
    "ppt/slideLayouts/slideLayout1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="10210800" cy="1371600"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></a:lstStyle><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:sldLayout>`,
  );
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    relationships([
      [
        "rId1",
        `${OFFICE_RELATIONSHIP}/slideMaster`,
        "../slideMasters/slideMaster1.xml",
      ],
    ]),
  );
  zip.file(
    "ppt/slideMasters/slideMaster1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="5" name="Decoration"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="6400800"/><a:ext cx="12192000" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></p:spPr>
    </p:sp>
  </p:spTree></p:cSld>
</p:sldMaster>`,
  );
  zip.file(
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    relationships([
      [
        "rId1",
        `${OFFICE_RELATIONSHIP}/theme`,
        "../theme/theme1.xml",
      ],
    ]),
  );
  zip.file(
    "ppt/theme/theme1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test">
  <a:themeElements><a:clrScheme name="Test">
    <a:dk1><a:srgbClr val="000000"/></a:dk1>
    <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="222222"/></a:dk2>
    <a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
    <a:accent1><a:srgbClr val="112233"/></a:accent1>
    <a:accent2><a:srgbClr val="CC5500"/></a:accent2>
  </a:clrScheme></a:themeElements>
</a:theme>`,
  );
  zip.file(
    "ppt/media/image1.png",
    Uint8Array.from(
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
        "base64",
      ),
    ),
  );
  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
  });
}

export async function createOleObjectPptx(
  options: { previewImage: boolean },
): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(await createMinimalPptx());
  const contentTypes = await zip.file("[Content_Types].xml")?.async("string");
  const slide = await zip.file("ppt/slides/slide1.xml")?.async("string");
  const slideRelationships = await zip
    .file("ppt/slides/_rels/slide1.xml.rels")
    ?.async("string");
  if (!contentTypes || !slide || !slideRelationships) {
    throw new Error("The minimal PPTX fixture is incomplete.");
  }
  const previewPicture = options.previewImage
    ? `<p:pic>
            <p:nvPicPr><p:cNvPr id="0" name="Embedded preview"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
            <p:blipFill><a:blip rel:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
            <p:spPr><a:xfrm><a:off x="914400" y="5029200"/><a:ext cx="1828800" cy="914400"/></a:xfrm></p:spPr>
          </p:pic>`
    : "";
  zip.file(
    "[Content_Types].xml",
    contentTypes.replace(
      '<Default Extension="png"',
      '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>\n  <Default Extension="png"',
    ),
  );
  zip.file(
    "ppt/slides/slide1.xml",
    slide.replace(
      "</p:spTree></p:cSld>",
      `<p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="5" name="Object"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="914400" y="5029200"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
        <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
          <mc:Choice Requires="v"><p:oleObj name="Photo" rel:id="rId4" progId="MSPhotoEd.3"><p:embed/></p:oleObj></mc:Choice>
          <mc:Fallback><p:oleObj name="Photo" rel:id="rId4" progId="MSPhotoEd.3"><p:embed/>${previewPicture}</p:oleObj></mc:Fallback>
        </mc:AlternateContent>
      </a:graphicData></a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>`,
    ),
  );
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    slideRelationships.replace(
      "</Relationships>",
      `<Relationship Id="rId4" Type="${OFFICE_RELATIONSHIP}/oleObject" Target="../embeddings/oleObject1.bin"/></Relationships>`,
    ),
  );
  zip.file(
    "ppt/embeddings/oleObject1.bin",
    Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0]),
  );
  return zip.generateAsync({ type: "arraybuffer" });
}

/**
 * A slide whose body placeholder inherits position, bullets, spacing and
 * colors from the layout and master, under an inverted layout color map.
 */
export async function createStyleInheritancePptx(): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(await createMinimalPptx());
  const namespaces = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:rel="${OFFICE_RELATIONSHIP}"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
  const groupProperties =
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';
  zip.file(
    "ppt/slideMasters/slideMaster1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<p:sldMaster ${namespaces}>
  <p:cSld><p:spTree>${groupProperties}
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="4525963"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Decoration"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="6400800"/><a:ext cx="12192000" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></p:spPr>
    </p:sp>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr marL="228600" indent="-228600"><a:spcBef><a:spcPts val="1000"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/><a:defRPr sz="2800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr>
      <a:lvl2pPr marL="685800" indent="-228600"><a:spcBef><a:spcPts val="1000"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="&#8211;"/><a:defRPr sz="2400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl2pPr>
    </p:bodyStyle>
    <p:otherStyle/>
  </p:txStyles>
</p:sldMaster>`,
  );
  zip.file(
    "ppt/slideLayouts/slideLayout1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<p:sldLayout ${namespaces} showMasterSp="0">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>${groupProperties}
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Content"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>
        <p:spPr/>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>
</p:sldLayout>`,
  );
  zip.file(
    "ppt/slides/slide1.xml",
    `<?xml version="1.0" encoding="UTF-8"?>
<p:sld ${namespaces}>
  <p:cSld><p:spTree>${groupProperties}
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Content"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:lstStyle/>
        <a:p><a:r><a:rPr lang="en-US"/><a:t>First level</a:t></a:r></a:p>
        <a:p><a:pPr lvl="1"/><a:r><a:rPr lang="en-US"/><a:t>Second level</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="3" name="Shrunk text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="6096000" y="457200"/><a:ext cx="3048000" cy="457200"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>
      <p:txBody><a:bodyPr><a:normAutofit fontScale="50000"/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="4000"/><a:t>Shrunk</a:t></a:r><a:r><a:rPr lang="en-US"/><a:t> / </a:t></a:r><a:fld type="slidenum"><a:rPr lang="en-US"/><a:t>#</a:t></a:fld></a:p></p:txBody>
    </p:sp>
    <p:cxnSp>
      <p:nvCxnSpPr><p:cNvPr id="4" name="Arrow"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
      <p:spPr><a:xfrm><a:off x="914400" y="6096000"/><a:ext cx="2743200" cy="0"/></a:xfrm><a:prstGeom prst="straightConnector1"/><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/><a:tailEnd type="triangle"/></a:ln></p:spPr>
    </p:cxnSp>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="5" name="Freeform"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="6096000" y="3048000"/><a:ext cx="1270000" cy="1270000"/></a:xfrm>
        <a:custGeom><a:avLst/><a:gdLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="50" y="0"/></a:lnTo><a:cubicBezTo><a:pt x="60" y="20"/><a:pt x="80" y="60"/><a:pt x="100" y="100"/></a:cubicBezTo><a:close/></a:path></a:pathLst></a:custGeom>
        <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00FF00"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"><a:alpha val="50000"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>
      </p:spPr>
    </p:sp>
  </p:spTree></p:cSld>
</p:sld>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
}

export async function createCompatibilityPptx(): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(await createMinimalPptx());
  const slide = await zip.file("ppt/slides/slide1.xml")?.async("string");
  if (!slide) {
    throw new Error("Minimal PPTX slide is missing.");
  }
  const compatibilityShapes = `
    <p:grpSp>
      <p:nvGrpSpPr><p:cNvPr id="10" name="Grouped shapes"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="6858000" y="3657600"/><a:ext cx="2743200" cy="1371600"/><a:chOff x="0" y="0"/><a:chExt cx="2743200" cy="1371600"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="11" name="Grouped ellipse"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="ellipse"/><a:solidFill><a:schemeClr val="accent1"><a:tint val="30000"/></a:schemeClr></a:solidFill></p:spPr>
        <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:r><a:rPr sz="1400"><a:latin typeface="Aptos"/></a:rPr><a:t>Group</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="12" name="Grouped arrow"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="1097280" y="228600"/><a:ext cx="1371600" cy="685800"/></a:xfrm><a:prstGeom prst="rightArrow"/><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></p:spPr>
      </p:sp>
    </p:grpSp>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="13" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="9601200" y="3657600"/><a:ext cx="1097280" cy="914400"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" rel:id="rId4"/></a:graphicData></a:graphic>
    </p:graphicFrame>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="14" name="SmartArt"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="9601200" y="4800600"/><a:ext cx="1097280" cy="914400"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" rel:dm="rId5"/></a:graphicData></a:graphic>
    </p:graphicFrame>`;
  zip.file(
    "ppt/slides/slide1.xml",
    slide.replace("</p:spTree>", `${compatibilityShapes}\n  </p:spTree>`),
  );
  return zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
  });
}

function relationships(
  entries: Array<[id: string, type: string, target: string]>,
): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">
${entries
    .map(
      ([id, type, target]) =>
        `  <Relationship Id="${id}" Type="${type}" Target="${target}"/>`,
    )
    .join("\n")}
</Relationships>`;
}
