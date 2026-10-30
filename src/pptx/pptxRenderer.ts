import {
  type PptxPackage,
  type PptxRelationship,
  type PptxSlideContext,
  resolvePackagePath,
} from "./pptxPackage";
import {
  type PptxPathSpec,
  buildGeometryPaths,
  isLinePreset,
} from "./pptxGeometry";
import { extractMetafileBitmap, isMetafileExtension } from "./pptxMetafile";
import {
  type PlaceholderIndex,
  type PptxParagraphStyle,
  type PptxRunStyle,
  findInheritedPlaceholders,
  formatAutoNumber,
  getPlaceholder,
  indexPlaceholders,
  masterTextStyle,
  resolveParagraphStyle,
  resolveRunStyle,
} from "./pptxTextStyle";
import {
  type PptxFill,
  type PptxTheme,
  fillToCssBackground,
  fillToCssColor,
  findColorMapOverride,
  findMasterColorMap,
  parsePptxTheme,
  readColorChoice,
  readFill,
  readOuterShadow,
  resolveFillReference,
} from "./pptxTheme";
import {
  attribute,
  childElements,
  childrenNamed,
  descendantsNamed,
  firstChildNamed,
  firstDescendantNamed,
  localName,
  namespacedAttribute,
  numberAttribute,
  textContent,
} from "./xml";
import { releaseResources } from "../reader/resources";

const NATURAL_SLIDE_WIDTH = 960;
const SVG_NAMESPACE = "http://www.w3.org/2000/svg";
const OFFICE_RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const EMU_PER_POINT = 12_700;
// PowerPoint's single line spacing is about 1.2 times the font size.
const SINGLE_LINE_HEIGHT = 1.2;
const DEFAULT_LINE_WIDTH_EMU = 9_525;
const SHAPE_ELEMENT_NAMES = ["sp", "pic", "graphicFrame", "cxnSp", "grpSp"];
const DASH_PATTERNS: Record<string, number[]> = {
  dot: [1, 3],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};
const ARROW_SIZES: Record<string, number> = { sm: 2, med: 3, lg: 5 };

// SVG definitions are referenced by document-wide IDs, and several slides can
// be mounted in the same document, so the counter is shared across renders.
let svgDefinitionCounter = 0;

export interface RenderedPptxSlide {
  element: HTMLElement;
  resources: Set<string>;
  width: number;
  height: number;
  diagnostics: PptxRenderDiagnostics;
}

export interface PptxRenderDiagnostics {
  durationMs: number;
  yieldCount: number;
  maxWorkSliceMs: number;
  layerCount: number;
  shapeCount: number;
  textShapeCount: number;
  imageCount: number;
  tableCount: number;
  chartCount: number;
  smartArtCount: number;
  unsupportedObjectCount: number;
  resourceCount: number;
  fontFamilies: string[];
}

export interface PptxRenderOptions {
  isCancelled?: () => boolean;
  now?: () => number;
  timeSliceMs?: number;
  yieldControl?: () => Promise<void>;
}

export class PptxRenderCancelledError extends Error {
  constructor() {
    super("PowerPoint rendering was cancelled.");
    this.name = "PptxRenderCancelledError";
  }
}

interface CoordinateSpace {
  offsetX: number;
  offsetY: number;
  scaleX: number;
  scaleY: number;
}

interface ShapeTransform {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipHorizontal: boolean;
  flipVertical: boolean;
}

interface LayerContext {
  kind: "master" | "layout" | "slide";
  xmlDocument: Document;
  path: string;
  relationships: Map<string, PptxRelationship>;
}

interface RenderContext {
  ownerDocument: Document;
  pptx: PptxPackage;
  theme: PptxTheme;
  master: Document | null;
  layoutPlaceholders: PlaceholderIndex;
  masterPlaceholders: PlaceholderIndex;
  slideNumber: number;
  resources: Set<string>;
  slideWidth: number;
  slideHeight: number;
  pixelWidth: number;
  pixelHeight: number;
  diagnostics: MutablePptxRenderDiagnostics;
  isCancelled: () => boolean;
  now: () => number;
  yieldControl: () => Promise<void>;
  timeSliceMs: number;
  lastYieldAt: number;
}

interface MutablePptxRenderDiagnostics {
  yieldCount: number;
  maxWorkSliceMs: number;
  layerCount: number;
  shapeCount: number;
  textShapeCount: number;
  imageCount: number;
  tableCount: number;
  chartCount: number;
  smartArtCount: number;
  unsupportedObjectCount: number;
  fontFamilies: Set<string>;
}

interface ResolvedLine {
  color: string;
  width: number;
  dash: string | null;
  headEnd: Element | null;
  tailEnd: Element | null;
}

interface TextScale {
  font: number;
  lineReduction: number;
}

export async function renderPptxSlide(
  pptx: PptxPackage,
  slideContext: PptxSlideContext,
  ownerDocument: Document,
  options: PptxRenderOptions = {},
): Promise<RenderedPptxSlide> {
  const now = options.now ?? (() => performance.now());
  const startedAt = now();
  const pixelWidth = NATURAL_SLIDE_WIDTH;
  const pixelHeight =
    NATURAL_SLIDE_WIDTH * (pptx.slideHeight / pptx.slideWidth);
  const stageEl = ownerDocument.createElement("div");
  stageEl.className = "pptx-reader-stage";
  stageEl.style.width = `${pixelWidth}px`;
  stageEl.style.height = `${pixelHeight}px`;
  const resources = new Set<string>();
  const diagnostics: MutablePptxRenderDiagnostics = {
    yieldCount: 0,
    maxWorkSliceMs: 0,
    layerCount: 0,
    shapeCount: 0,
    textShapeCount: 0,
    imageCount: 0,
    tableCount: 0,
    chartCount: 0,
    smartArtCount: 0,
    unsupportedObjectCount: 0,
    fontFamilies: new Set<string>(),
  };
  const theme = parsePptxTheme(slideContext.theme, [
    findMasterColorMap(slideContext.master),
    findColorMapOverride(slideContext.layout),
    findColorMapOverride(slideContext.slide),
  ]);
  const context: RenderContext = {
    ownerDocument,
    pptx,
    theme,
    master: slideContext.master,
    layoutPlaceholders: indexPlaceholders(slideContext.layout),
    masterPlaceholders: indexPlaceholders(slideContext.master),
    slideNumber: pptx.slidePaths.indexOf(slideContext.slidePath) + 1,
    resources,
    slideWidth: pptx.slideWidth,
    slideHeight: pptx.slideHeight,
    pixelWidth,
    pixelHeight,
    diagnostics,
    isCancelled: options.isCancelled ?? (() => false),
    now,
    yieldControl:
      options.yieldControl ?? createRenderYield(ownerDocument),
    timeSliceMs: Math.max(1, options.timeSliceMs ?? 8),
    lastYieldAt: startedAt,
  };

  try {
    ensureRenderActive(context);
    stageEl.style.color = theme.colors.tx1 ?? "#000000";
    await renderBackground(stageEl, slideContext, context);
    const rootSpace: CoordinateSpace = {
      offsetX: 0,
      offsetY: 0,
      scaleX: 1,
      scaleY: 1,
    };

    const slideHidesMaster = hidesMasterShapes(slideContext.slide);
    const layoutHidesMaster = hidesMasterShapes(slideContext.layout);
    if (
      slideContext.master &&
      slideContext.masterPath &&
      !slideHidesMaster &&
      !layoutHidesMaster
    ) {
      await renderLayer(
        stageEl,
        {
          kind: "master",
          xmlDocument: slideContext.master,
          path: slideContext.masterPath,
          relationships: slideContext.masterRelationships,
        },
        context,
        rootSpace,
      );
    }
    if (slideContext.layout && slideContext.layoutPath && !slideHidesMaster) {
      await renderLayer(
        stageEl,
        {
          kind: "layout",
          xmlDocument: slideContext.layout,
          path: slideContext.layoutPath,
          relationships: slideContext.layoutRelationships,
        },
        context,
        rootSpace,
      );
    }
    await renderLayer(
      stageEl,
      {
        kind: "slide",
        xmlDocument: slideContext.slide,
        path: slideContext.slidePath,
        relationships: slideContext.slideRelationships,
      },
      context,
      rootSpace,
    );
    ensureRenderActive(context);
    recordWorkSlice(context);

    return {
      element: stageEl,
      resources,
      width: pixelWidth,
      height: pixelHeight,
      diagnostics: {
        durationMs: Math.max(0, now() - startedAt),
        yieldCount: diagnostics.yieldCount,
        maxWorkSliceMs: diagnostics.maxWorkSliceMs,
        layerCount: diagnostics.layerCount,
        shapeCount: diagnostics.shapeCount,
        textShapeCount: diagnostics.textShapeCount,
        imageCount: diagnostics.imageCount,
        tableCount: diagnostics.tableCount,
        chartCount: diagnostics.chartCount,
        smartArtCount: diagnostics.smartArtCount,
        unsupportedObjectCount: diagnostics.unsupportedObjectCount,
        resourceCount: resources.size,
        fontFamilies: [...diagnostics.fontFamilies].sort(),
      },
    };
  } catch (error) {
    releaseResources(resources, (resource) => {
      URL.revokeObjectURL(resource);
    });
    throw error;
  }
}

async function renderLayer(
  stageEl: HTMLElement,
  layer: LayerContext,
  context: RenderContext,
  coordinateSpace: CoordinateSpace,
): Promise<void> {
  ensureRenderActive(context);
  const shapeTree = firstDescendantNamed(
    layer.xmlDocument.documentElement,
    "spTree",
  );
  if (!shapeTree) {
    return;
  }
  context.diagnostics.layerCount += 1;
  for (const shape of childElements(shapeTree)) {
    ensureRenderActive(context);
    await yieldIfNeeded(context);
    if (!SHAPE_ELEMENT_NAMES.includes(localName(shape))) {
      continue;
    }
    // Master and layout placeholders only provide defaults for slide shapes.
    if (layer.kind !== "slide" && getPlaceholder(shape)) {
      continue;
    }
    await renderShape(
      stageEl,
      shape,
      layer,
      context,
      coordinateSpace,
    );
  }
}

async function renderShape(
  parentEl: HTMLElement,
  shape: Element,
  layer: LayerContext,
  context: RenderContext,
  coordinateSpace: CoordinateSpace,
): Promise<void> {
  ensureRenderActive(context);
  context.diagnostics.shapeCount += 1;
  if (localName(shape) === "grpSp") {
    const groupTransform = parseGroupTransform(shape, coordinateSpace);
    if (!groupTransform) {
      return;
    }
    for (const child of childElements(shape)) {
      await yieldIfNeeded(context);
      if (SHAPE_ELEMENT_NAMES.includes(localName(child))) {
        await renderShape(parentEl, child, layer, context, groupTransform);
      }
    }
    return;
  }

  const inherited =
    layer.kind === "slide"
      ? findInheritedPlaceholders(
          shape,
          context.layoutPlaceholders,
          context.masterPlaceholders,
        )
      : [];
  const transform = parseShapeTransform(shape, inherited, coordinateSpace);
  if (!transform || transform.width < 0 || transform.height < 0) {
    return;
  }
  const lineLike = isLineShape(shape, inherited);
  const hasArea = transform.width > 0 && transform.height > 0;
  const hasLength = transform.width > 0 || transform.height > 0;
  if (!hasArea && !(lineLike && hasLength)) {
    return;
  }
  const wrapperEl = context.ownerDocument.createElement("div");
  wrapperEl.className = `pptx-reader-element pptx-reader-${localName(shape)}`;
  applyTransform(wrapperEl, transform, context);
  parentEl.appendChild(wrapperEl);

  switch (localName(shape)) {
    case "pic":
      context.diagnostics.imageCount += 1;
      await renderPicture(wrapperEl, shape, layer, context);
      break;
    case "graphicFrame":
      await renderGraphicFrame(wrapperEl, shape, layer, context);
      break;
    case "cxnSp":
    case "sp":
    default:
      await renderShapeGeometry(
        wrapperEl,
        shape,
        inherited,
        transform,
        layer,
        context,
      );
      await renderShapeText(wrapperEl, shape, inherited, context);
      break;
  }
}

async function renderShapeGeometry(
  wrapperEl: HTMLElement,
  shape: Element,
  inherited: Element[],
  transform: ShapeTransform,
  layer: LayerContext,
  context: RenderContext,
): Promise<void> {
  const propertyChain = [shape, ...inherited]
    .map(getShapeProperties)
    .filter((properties): properties is Element => properties !== null);
  const geometryProperties =
    propertyChain.find(
      (properties) =>
        firstChildNamed(properties, "custGeom") ??
        firstChildNamed(properties, "prstGeom"),
    ) ?? null;
  const preset =
    attribute(firstChildNamed(geometryProperties, "prstGeom"), "prst") ?? "rect";
  const lineOnly =
    !firstChildNamed(geometryProperties, "custGeom") && isLinePreset(preset);
  const style = firstChildNamed(shape, "style");

  let fill: PptxFill | null = null;
  if (!lineOnly) {
    for (const properties of propertyChain) {
      fill = readFill(properties, context.theme);
      if (fill) {
        break;
      }
    }
    fill ??= resolveFillReference(
      firstChildNamed(style, "fillRef"),
      context.theme,
    );
  }
  if (fill?.type === "image") {
    await renderImage(wrapperEl, fill.relationshipId, null, layer, context, "");
    fill = null;
  }
  const line = resolveLine(
    propertyChain,
    style,
    localName(shape) === "cxnSp",
    context,
  );
  if ((!fill || fill.type === "none") && !line) {
    return;
  }

  const width = toPixelsX(transform.width, context);
  const height = toPixelsY(transform.height, context);
  const paths = buildGeometryPaths(
    geometryProperties,
    width,
    height,
    transform.width,
    transform.height,
  );
  const svgEl = context.ownerDocument.createElementNS(SVG_NAMESPACE, "svg");
  svgEl.classList.add("pptx-reader-shape-svg");
  svgEl.setAttribute("width", String(Math.max(1, width)));
  svgEl.setAttribute("height", String(Math.max(1, height)));
  svgEl.setAttribute("overflow", "visible");
  const definitionsEl = context.ownerDocument.createElementNS(SVG_NAMESPACE, "defs");
  let hasDefinitions = false;

  let fillPaint = "none";
  if (fill?.type === "solid") {
    fillPaint = fill.color;
  } else if (fill?.type === "gradient") {
    fillPaint = `url(#${appendGradient(definitionsEl, fill, context)})`;
    hasDefinitions = true;
  }
  let markerStart: string | null = null;
  let markerEnd: string | null = null;
  if (line) {
    markerStart = appendArrowMarker(definitionsEl, line.headEnd, line.color, context);
    markerEnd = appendArrowMarker(definitionsEl, line.tailEnd, line.color, context);
    hasDefinitions ||= markerStart !== null || markerEnd !== null;
  }
  if (hasDefinitions) {
    svgEl.appendChild(definitionsEl);
  }

  // Arrowheads belong to the last open, stroked path of the geometry.
  const markedPath = paths
    .filter((path) => path.stroke && (lineOnly || !path.fill))
    .at(-1);
  for (const path of paths) {
    const pathEl = createPathElement(
      path,
      lineOnly ? "none" : fillPaint,
      line,
      context,
    );
    if (path === markedPath && markerStart) {
      pathEl.setAttribute("marker-start", `url(#${markerStart})`);
    }
    if (path === markedPath && markerEnd) {
      pathEl.setAttribute("marker-end", `url(#${markerEnd})`);
    }
    svgEl.appendChild(pathEl);
  }

  const effectList =
    propertyChain
      .map((properties) => firstChildNamed(properties, "effectLst"))
      .find((effects) => effects !== null) ?? null;
  const shadow = readOuterShadow(effectList, context.theme);
  if (shadow) {
    svgEl.style.filter = `drop-shadow(${emuToPixels(shadow.offsetX, context)}px ${emuToPixels(shadow.offsetY, context)}px ${emuToPixels(shadow.blur, context) / 2}px ${shadow.color})`;
  }
  wrapperEl.appendChild(svgEl);
}

function createPathElement(
  path: PptxPathSpec,
  fillPaint: string,
  line: ResolvedLine | null,
  context: RenderContext,
): SVGElement {
  const pathEl = context.ownerDocument.createElementNS(SVG_NAMESPACE, "path");
  pathEl.setAttribute("d", path.d);
  pathEl.setAttribute("fill", path.fill ? fillPaint : "none");
  if (line && path.stroke) {
    pathEl.setAttribute("stroke", line.color);
    pathEl.setAttribute("stroke-width", String(line.width));
    pathEl.setAttribute("stroke-linejoin", "round");
    if (line.dash) {
      pathEl.setAttribute("stroke-dasharray", line.dash);
    }
  } else {
    pathEl.setAttribute("stroke", "none");
  }
  return pathEl;
}

function resolveLine(
  propertyChain: Element[],
  style: Element | null,
  isConnector: boolean,
  context: RenderContext,
): ResolvedLine | null {
  const lines = propertyChain
    .map((properties) => firstChildNamed(properties, "ln"))
    .filter((line): line is Element => line !== null);
  const lineReference = firstChildNamed(style, "lnRef");
  const referenceIndex = numberAttribute(lineReference, "idx", 0);
  const themeLine =
    referenceIndex > 0 ? context.theme.lineStyles[referenceIndex - 1] ?? null : null;
  const referenceColor = readColorChoice(lineReference, context.theme);

  let fill: PptxFill | null = null;
  for (const line of lines) {
    fill = readFill(line, context.theme);
    if (fill) {
      break;
    }
  }
  if (!fill && referenceColor && referenceIndex > 0) {
    fill =
      (themeLine ? readFill(themeLine, context.theme, referenceColor) : null) ??
      { type: "solid", color: referenceColor };
  }
  if (!fill && isConnector && !lineReference) {
    fill = { type: "solid", color: context.theme.colors.tx1 ?? "#000000" };
  }
  const color = fillToCssColor(fill);
  if (!fill || fill.type === "none" || !color) {
    return null;
  }

  const widthSource = [...lines, ...(themeLine ? [themeLine] : [])].find(
    (line) => attribute(line, "w") !== null,
  );
  const width = Math.max(
    0.5,
    emuToPixels(numberAttribute(widthSource, "w", DEFAULT_LINE_WIDTH_EMU), context),
  );
  const dashName = [...lines, ...(themeLine ? [themeLine] : [])]
    .map((line) => attribute(firstChildNamed(line, "prstDash"), "val"))
    .find((value) => value !== null);
  const dashPattern = dashName ? DASH_PATTERNS[dashName] : undefined;
  return {
    color,
    width,
    dash: dashPattern
      ? dashPattern.map((segment) => segment * width).join(" ")
      : null,
    headEnd: lines.map((line) => firstChildNamed(line, "headEnd")).find(Boolean) ?? null,
    tailEnd: lines.map((line) => firstChildNamed(line, "tailEnd")).find(Boolean) ?? null,
  };
}

function appendGradient(
  definitionsEl: SVGElement,
  fill: Extract<PptxFill, { type: "gradient" }>,
  context: RenderContext,
): string {
  svgDefinitionCounter += 1;
  const id = `pptx-reader-gradient-${svgDefinitionCounter}`;
  const gradientEl = context.ownerDocument.createElementNS(
    SVG_NAMESPACE,
    fill.radial ? "radialGradient" : "linearGradient",
  );
  gradientEl.setAttribute("id", id);
  if (!fill.radial) {
    const radians = (fill.angle * Math.PI) / 180;
    const dx = Math.cos(radians) / 2;
    const dy = Math.sin(radians) / 2;
    gradientEl.setAttribute("x1", String(0.5 - dx));
    gradientEl.setAttribute("y1", String(0.5 - dy));
    gradientEl.setAttribute("x2", String(0.5 + dx));
    gradientEl.setAttribute("y2", String(0.5 + dy));
  }
  for (const stop of fill.stops) {
    const stopEl = context.ownerDocument.createElementNS(SVG_NAMESPACE, "stop");
    stopEl.setAttribute("offset", String(stop.offset));
    stopEl.setAttribute("stop-color", stop.color);
    gradientEl.appendChild(stopEl);
  }
  definitionsEl.appendChild(gradientEl);
  return id;
}

function appendArrowMarker(
  definitionsEl: SVGElement,
  end: Element | null,
  color: string,
  context: RenderContext,
): string | null {
  const type = attribute(end, "type");
  if (!type || type === "none") {
    return null;
  }
  svgDefinitionCounter += 1;
  const id = `pptx-reader-marker-${svgDefinitionCounter}`;
  const markerEl = context.ownerDocument.createElementNS(SVG_NAMESPACE, "marker");
  const markerWidth = ARROW_SIZES[attribute(end, "len") ?? "med"] ?? 3;
  const markerHeight = ARROW_SIZES[attribute(end, "w") ?? "med"] ?? 3;
  markerEl.setAttribute("id", id);
  markerEl.setAttribute("viewBox", "0 0 10 10");
  markerEl.setAttribute("markerUnits", "strokeWidth");
  markerEl.setAttribute("markerWidth", String(markerWidth));
  markerEl.setAttribute("markerHeight", String(markerHeight));
  markerEl.setAttribute("preserveAspectRatio", "none");
  markerEl.setAttribute("orient", "auto-start-reverse");
  markerEl.setAttribute("refY", "5");
  const shapeEl = context.ownerDocument.createElementNS(SVG_NAMESPACE, "path");
  switch (type) {
    case "oval":
      markerEl.setAttribute("refX", "5");
      shapeEl.setAttribute("d", "M0,5 A5,5 0 1 1 10,5 A5,5 0 1 1 0,5 Z");
      shapeEl.setAttribute("fill", color);
      break;
    case "diamond":
      markerEl.setAttribute("refX", "5");
      shapeEl.setAttribute("d", "M0,5 L5,0 L10,5 L5,10 Z");
      shapeEl.setAttribute("fill", color);
      break;
    case "arrow":
      markerEl.setAttribute("refX", "9");
      shapeEl.setAttribute("d", "M0,0 L10,5 L0,10");
      shapeEl.setAttribute("fill", "none");
      shapeEl.setAttribute("stroke", color);
      shapeEl.setAttribute("stroke-width", String(10 / markerHeight));
      break;
    case "stealth":
      markerEl.setAttribute("refX", "10");
      shapeEl.setAttribute("d", "M0,0 L10,5 L0,10 L3,5 Z");
      shapeEl.setAttribute("fill", color);
      break;
    default:
      markerEl.setAttribute("refX", "10");
      shapeEl.setAttribute("d", "M0,0 L10,5 L0,10 Z");
      shapeEl.setAttribute("fill", color);
      break;
  }
  markerEl.appendChild(shapeEl);
  definitionsEl.appendChild(markerEl);
  return id;
}

async function renderShapeText(
  wrapperEl: HTMLElement,
  shape: Element,
  inherited: Element[],
  context: RenderContext,
): Promise<void> {
  const textBody = firstChildNamed(shape, "txBody");
  if (
    !textBody ||
    (descendantsNamed(textBody, "t").length === 0 &&
      descendantsNamed(textBody, "fld").length === 0)
  ) {
    return;
  }
  context.diagnostics.textShapeCount += 1;

  const inheritedBodies = inherited
    .map((candidate) => firstChildNamed(candidate, "txBody"))
    .filter((body): body is Element => body !== null);
  const listStyles = [textBody, ...inheritedBodies]
    .map((body) => firstChildNamed(body, "lstStyle"))
    .filter((listStyle): listStyle is Element => listStyle !== null);
  const masterStyle = masterTextStyle(shape, context.master);
  if (masterStyle) {
    listStyles.push(masterStyle);
  }
  if (context.pptx.defaultTextStyle) {
    listStyles.push(context.pptx.defaultTextStyle);
  }

  const bodyProperties = [textBody, ...inheritedBodies]
    .map((body) => firstChildNamed(body, "bodyPr"))
    .filter((properties): properties is Element => properties !== null);
  const bodyAttribute = (name: string): string | null => {
    for (const properties of bodyProperties) {
      const value = attribute(properties, name);
      if (value !== null) {
        return value;
      }
    }
    return null;
  };
  const inset = (name: string, fallback: number): number =>
    emuToPixels(Number(bodyAttribute(name) ?? fallback) || 0, context);
  const autofit = firstChildNamed(firstChildNamed(textBody, "bodyPr"), "normAutofit");
  const scale: TextScale = {
    font: numberAttribute(autofit, "fontScale", 100_000) / 100_000,
    lineReduction: numberAttribute(autofit, "lnSpcReduction", 0) / 100_000,
  };

  const textEl = context.ownerDocument.createElement("div");
  textEl.className = "pptx-reader-text";
  const anchor = bodyAttribute("anchor");
  textEl.style.justifyContent =
    anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : "flex-start";
  textEl.style.padding = `${inset("tIns", 45_720)}px ${inset("rIns", 91_440)}px ${inset("bIns", 45_720)}px ${inset("lIns", 91_440)}px`;
  if (bodyAttribute("wrap") === "none") {
    textEl.classList.add("is-unwrapped");
  }
  const vertical = bodyAttribute("vert");
  if (vertical && vertical !== "horz") {
    textEl.classList.add(vertical === "vert270" ? "is-vertical-270" : "is-vertical");
  }

  const fontReference = firstChildNamed(firstChildNamed(shape, "style"), "fontRef");
  const numbering = new Map<number, number>();
  const paragraphs = childrenNamed(textBody, "p");
  for (const [index, paragraph] of paragraphs.entries()) {
    await yieldIfNeeded(context);
    textEl.appendChild(
      await renderParagraph(
        paragraph,
        listStyles,
        fontReference,
        scale,
        numbering,
        index === 0,
        context,
      ),
    );
  }
  wrapperEl.appendChild(textEl);
}

async function renderParagraph(
  paragraph: Element,
  listStyles: Element[],
  fontReference: Element | null,
  scale: TextScale,
  numbering: Map<number, number>,
  isFirst: boolean,
  context: RenderContext,
): Promise<HTMLElement> {
  const style = resolveParagraphStyle(paragraph, listStyles, context.theme);
  const paragraphEl = context.ownerDocument.createElement("div");
  paragraphEl.className = "pptx-reader-paragraph";
  paragraphEl.style.textAlign =
    style.alignment === "ctr"
      ? "center"
      : style.alignment === "r"
        ? "right"
        : style.alignment === "just" || style.alignment === "dist"
          ? "justify"
          : "left";

  const runs = childElements(paragraph).filter((run) =>
    ["r", "br", "fld"].includes(localName(run)),
  );
  const textRuns = runs.filter((run) => localName(run) !== "br");
  const hasText = textRuns.some((run) => runText(run, context) !== "");
  const firstRunProperties =
    firstChildNamed(textRuns[0] ?? null, "rPr") ??
    firstChildNamed(paragraph, "endParaRPr");
  const baseRun = resolveRunStyle(
    firstRunProperties,
    style,
    fontReference,
    context.theme,
  );
  const baseSize = pointsToPixels(baseRun.sizePoints * scale.font, context);
  paragraphEl.style.fontSize = `${baseSize}px`;
  paragraphEl.style.lineHeight = lineHeight(style, scale, context);
  if (style.marginLeft) {
    paragraphEl.style.paddingLeft = `${emuToPixels(style.marginLeft, context)}px`;
  }
  if (style.indent) {
    paragraphEl.style.textIndent = `${emuToPixels(style.indent, context)}px`;
  }
  // PowerPoint ignores space before the first paragraph of a text body.
  if (!isFirst && style.spaceBefore) {
    paragraphEl.style.marginTop = `${spacingToPixels(style.spaceBefore, baseSize, context)}px`;
  }
  if (style.spaceAfter) {
    paragraphEl.style.marginBottom = `${spacingToPixels(style.spaceAfter, baseSize, context)}px`;
  }

  for (const level of [...numbering.keys()]) {
    if (level > style.level) {
      numbering.delete(level);
    }
  }
  if (style.bullet?.kind !== "number" && hasText) {
    numbering.delete(style.level);
  }
  if (style.bullet && hasText) {
    let label = style.bullet.character;
    if (style.bullet.kind === "number") {
      const value = numbering.get(style.level) ?? style.bullet.startAt;
      numbering.set(style.level, value + 1);
      label = formatAutoNumber(style.bullet.numberScheme, value);
    }
    const bulletEl = context.ownerDocument.createElement("span");
    bulletEl.className = "pptx-reader-bullet";
    bulletEl.textContent = label;
    const bulletSize = style.bullet.sizePoints
      ? pointsToPixels(style.bullet.sizePoints * scale.font, context)
      : baseSize * (style.bullet.sizePercent ?? 1);
    bulletEl.style.fontSize = `${bulletSize}px`;
    const bulletColor = style.bullet.color ?? baseRun.color;
    if (bulletColor) {
      bulletEl.style.color = bulletColor;
    }
    bulletEl.style.fontFamily = fontFamily(
      style.bullet.typeface ? [style.bullet.typeface] : baseRun.typefaces,
    );
    const hangingIndent = -emuToPixels(style.indent, context);
    if (hangingIndent > 0) {
      bulletEl.classList.add("is-hanging");
      bulletEl.style.width = `${hangingIndent}px`;
    }
    paragraphEl.appendChild(bulletEl);
  }

  for (const run of runs) {
    await yieldIfNeeded(context);
    if (localName(run) === "br") {
      paragraphEl.appendChild(context.ownerDocument.createElement("br"));
      continue;
    }
    const value = runText(run, context);
    if (!value) {
      continue;
    }
    const spanEl = context.ownerDocument.createElement("span");
    spanEl.textContent = value;
    applyRunStyle(
      spanEl,
      resolveRunStyle(
        firstChildNamed(run, "rPr"),
        style,
        fontReference,
        context.theme,
      ),
      scale,
      context,
    );
    paragraphEl.appendChild(spanEl);
  }
  if (!paragraphEl.hasChildNodes()) {
    paragraphEl.appendChild(context.ownerDocument.createElement("br"));
  }
  return paragraphEl;
}

function runText(run: Element, context: RenderContext): string {
  if (localName(run) === "fld" && attribute(run, "type") === "slidenum") {
    return context.slideNumber > 0 ? String(context.slideNumber) : textContent(firstChildNamed(run, "t"));
  }
  return textContent(firstChildNamed(run, "t"));
}

function lineHeight(
  style: PptxParagraphStyle,
  scale: TextScale,
  context: RenderContext,
): string {
  const spacing = style.lineSpacing;
  if (spacing?.type === "points") {
    return `${pointsToPixels(spacing.value, context)}px`;
  }
  const percent = Math.max(0.1, (spacing?.value ?? 1) - scale.lineReduction);
  return String(percent * SINGLE_LINE_HEIGHT);
}

function spacingToPixels(
  spacing: NonNullable<PptxParagraphStyle["spaceBefore"]>,
  fontSize: number,
  context: RenderContext,
): number {
  return spacing.type === "points"
    ? pointsToPixels(spacing.value, context)
    : spacing.value * fontSize * SINGLE_LINE_HEIGHT;
}

function applyRunStyle(
  spanEl: HTMLSpanElement,
  style: PptxRunStyle,
  scale: TextScale,
  context: RenderContext,
): void {
  let size = pointsToPixels(style.sizePoints * scale.font, context);
  if (style.baseline !== 0) {
    size *= 0.67;
    spanEl.style.verticalAlign = style.baseline > 0 ? "super" : "sub";
  }
  spanEl.style.fontSize = `${size}px`;
  if (style.bold) {
    spanEl.classList.add("is-bold");
  }
  if (style.italic) {
    spanEl.classList.add("is-italic");
  }
  if (style.underline) {
    spanEl.classList.add("is-underlined");
  }
  if (style.strike) {
    spanEl.classList.add("is-strikethrough");
  }
  if (style.capitals === "all") {
    spanEl.classList.add("is-all-caps");
  } else if (style.capitals === "small") {
    spanEl.classList.add("is-small-caps");
  }
  if (style.spacingPoints) {
    spanEl.style.letterSpacing = `${pointsToPixels(style.spacingPoints, context)}px`;
  }
  if (style.typefaces.length > 0) {
    spanEl.style.fontFamily = fontFamily(style.typefaces);
    for (const typeface of style.typefaces) {
      context.diagnostics.fontFamilies.add(typeface);
    }
  }
  if (style.color) {
    spanEl.style.color = style.color;
  }
  if (style.highlight) {
    spanEl.style.backgroundColor = style.highlight;
  }
}

// Missing theme fonts fall back to the metric-compatible faces PowerPoint
// substitutes on desktop systems before the platform default.
const FALLBACK_FONT_FAMILIES = ["Arial", "Helvetica", "sans-serif"];

function fontFamily(typefaces: string[]): string {
  return [
    ...typefaces.map((typeface) => `"${typeface.replace(/"/g, "")}"`),
    ...FALLBACK_FONT_FAMILIES,
  ].join(", ");
}

async function renderPicture(
  wrapperEl: HTMLElement,
  shape: Element,
  layer: LayerContext,
  context: RenderContext,
): Promise<void> {
  ensureRenderActive(context);
  const blipFill = firstDescendantNamed(shape, "blipFill");
  const blip = firstChildNamed(blipFill, "blip") ?? firstDescendantNamed(shape, "blip");
  const relationshipId = namespacedAttribute(
    blip,
    "embed",
    OFFICE_RELATIONSHIP_NAMESPACE,
  );
  const cropped = await renderImage(
    wrapperEl,
    relationshipId,
    firstChildNamed(blipFill, "srcRect"),
    layer,
    context,
    attribute(firstDescendantNamed(shape, "cNvPr"), "descr") ??
      attribute(firstDescendantNamed(shape, "cNvPr"), "name") ??
      "",
  );
  if (cropped) {
    wrapperEl.classList.add("is-cropped");
  }
}

/** Renders a package image. Returns true when a source crop was applied. */
async function renderImage(
  wrapperEl: HTMLElement,
  relationshipId: string | null,
  sourceRectangle: Element | null,
  layer: LayerContext,
  context: RenderContext,
  altText: string,
): Promise<boolean> {
  const relationship = relationshipId
    ? layer.relationships.get(relationshipId)
    : undefined;
  if (!relationship || relationship.external) {
    renderUnsupportedPlaceholder(wrapperEl, context, "Image");
    return false;
  }
  const imagePath = resolvePackagePath(layer.path, relationship.target);
  const extension = imagePath.split(".").at(-1)?.toLowerCase() ?? "";
  // Check the type before reading so non-image parts are never decompressed.
  const mimeType = context.pptx.getImageMimeType(imagePath);
  if (!mimeType && !isMetafileExtension(extension)) {
    renderUnsupportedPlaceholder(
      wrapperEl,
      context,
      "Unsupported image",
    );
    return false;
  }
  const bytes = await context.pptx.getBinary(imagePath);
  ensureRenderActive(context);
  if (!bytes) {
    renderUnsupportedPlaceholder(wrapperEl, context, "Image");
    return false;
  }
  const image = mimeType
    ? { bytes, mimeType }
    : extractMetafileBitmap(bytes, extension);
  if (!image) {
    renderUnsupportedPlaceholder(
      wrapperEl,
      context,
      "Unsupported image",
    );
    return false;
  }
  const blob = createBlob(image.bytes, image.mimeType);
  const url = URL.createObjectURL(blob);
  context.resources.add(url);
  const imageEl = context.ownerDocument.createElement("img");
  imageEl.className = "pptx-reader-image";
  imageEl.src = url;
  imageEl.alt = altText;
  imageEl.loading = "lazy";
  imageEl.decoding = "async";

  const left = numberAttribute(sourceRectangle, "l", 0) / 100_000;
  const top = numberAttribute(sourceRectangle, "t", 0) / 100_000;
  const right = numberAttribute(sourceRectangle, "r", 0) / 100_000;
  const bottom = numberAttribute(sourceRectangle, "b", 0) / 100_000;
  const visibleWidth = 1 - left - right;
  const visibleHeight = 1 - top - bottom;
  const cropped =
    [left, top, right, bottom].some((edge) => edge !== 0) &&
    visibleWidth > 0 &&
    visibleHeight > 0;
  if (cropped) {
    imageEl.classList.add("is-cropped");
    imageEl.style.width = `${100 / visibleWidth}%`;
    imageEl.style.height = `${100 / visibleHeight}%`;
    imageEl.style.left = `${(-left / visibleWidth) * 100}%`;
    imageEl.style.top = `${(-top / visibleHeight) * 100}%`;
  }
  wrapperEl.appendChild(imageEl);
  return cropped;
}

async function renderBackground(
  stageEl: HTMLElement,
  slideContext: PptxSlideContext,
  context: RenderContext,
): Promise<void> {
  const layers: Array<LayerContext | null> = [
    {
      kind: "slide",
      xmlDocument: slideContext.slide,
      path: slideContext.slidePath,
      relationships: slideContext.slideRelationships,
    },
    slideContext.layout && slideContext.layoutPath
      ? {
          kind: "layout",
          xmlDocument: slideContext.layout,
          path: slideContext.layoutPath,
          relationships: slideContext.layoutRelationships,
        }
      : null,
    slideContext.master && slideContext.masterPath
      ? {
          kind: "master",
          xmlDocument: slideContext.master,
          path: slideContext.masterPath,
          relationships: slideContext.masterRelationships,
        }
      : null,
  ];
  for (const layer of layers) {
    if (!layer) {
      continue;
    }
    const background = firstDescendantNamed(layer.xmlDocument.documentElement, "bg");
    if (!background) {
      continue;
    }
    const fill =
      readFill(firstChildNamed(background, "bgPr"), context.theme) ??
      resolveFillReference(firstChildNamed(background, "bgRef"), context.theme);
    if (!fill) {
      continue;
    }
    if (fill.type === "image") {
      stageEl.style.background = context.theme.colors.bg1 ?? "#ffffff";
      const backgroundEl = context.ownerDocument.createElement("div");
      backgroundEl.className = "pptx-reader-background";
      await renderImage(backgroundEl, fill.relationshipId, null, layer, context, "");
      stageEl.appendChild(backgroundEl);
      return;
    }
    const css = fillToCssBackground(fill);
    if (css) {
      stageEl.style.background = css;
      return;
    }
  }
  stageEl.style.background = context.theme.colors.bg1 ?? "#ffffff";
}

async function renderGraphicFrame(
  wrapperEl: HTMLElement,
  shape: Element,
  layer: LayerContext,
  context: RenderContext,
): Promise<void> {
  const table = firstDescendantNamed(shape, "tbl");
  if (!table) {
    const objectType = getGraphicFrameObjectType(shape);
    if (objectType === "Embedded object") {
      // The embedded OLE payload is never read; only the preview image that
      // PowerPoint stores beside it is rendered.
      const previewPicture = firstDescendantNamed(shape, "pic");
      if (previewPicture) {
        context.diagnostics.imageCount += 1;
        await renderPicture(wrapperEl, previewPicture, layer, context);
        return;
      }
    }
    if (objectType === "Chart") {
      context.diagnostics.chartCount += 1;
    } else if (objectType === "SmartArt") {
      context.diagnostics.smartArtCount += 1;
    }
    renderUnsupportedPlaceholder(
      wrapperEl,
      context,
      objectType,
    );
    return;
  }
  context.diagnostics.tableCount += 1;

  const tableEl = context.ownerDocument.createElement("table");
  tableEl.className = "pptx-reader-table";
  const gridColumns = descendantsNamed(
    firstChildNamed(table, "tblGrid") ?? table,
    "gridCol",
  ).map((column) => numberAttribute(column, "w", 1));
  const totalGridWidth = gridColumns.reduce((total, width) => total + width, 0);
  if (totalGridWidth > 0) {
    const columnGroupEl = context.ownerDocument.createElement("colgroup");
    for (const width of gridColumns) {
      await yieldIfNeeded(context);
      const columnEl = context.ownerDocument.createElement("col");
      columnEl.style.width = `${(width / totalGridWidth) * 100}%`;
      columnGroupEl.appendChild(columnEl);
    }
    tableEl.appendChild(columnGroupEl);
  }

  const listStyles = context.pptx.defaultTextStyle
    ? [context.pptx.defaultTextStyle]
    : [];
  for (const row of childrenNamed(table, "tr")) {
    await yieldIfNeeded(context);
    const rowEl = context.ownerDocument.createElement("tr");
    const rowHeight = numberAttribute(row, "h", 0);
    if (rowHeight > 0) {
      rowEl.style.height = `${emuToPixels(rowHeight, context)}px`;
    }
    for (const cell of childrenNamed(row, "tc")) {
      await yieldIfNeeded(context);
      // Cells covered by a merge are represented by the merge origin.
      if (attribute(cell, "hMerge") === "1" || attribute(cell, "vMerge") === "1") {
        continue;
      }
      const cellEl = context.ownerDocument.createElement("td");
      const gridSpan = numberAttribute(cell, "gridSpan", 1);
      const rowSpan = numberAttribute(cell, "rowSpan", 1);
      if (gridSpan > 1) {
        cellEl.colSpan = gridSpan;
      }
      if (rowSpan > 1) {
        cellEl.rowSpan = rowSpan;
      }
      const cellProperties = firstChildNamed(cell, "tcPr");
      applyCellStyle(cellEl, cellProperties, context);
      const textBody = firstChildNamed(cell, "txBody");
      if (textBody) {
        const cellListStyles = [
          ...[firstChildNamed(textBody, "lstStyle")].filter(
            (listStyle): listStyle is Element => listStyle !== null,
          ),
          ...listStyles,
        ];
        const numbering = new Map<number, number>();
        for (const [index, paragraph] of childrenNamed(textBody, "p").entries()) {
          await yieldIfNeeded(context);
          cellEl.appendChild(
            await renderParagraph(
              paragraph,
              cellListStyles,
              null,
              { font: 1, lineReduction: 0 },
              numbering,
              index === 0,
              context,
            ),
          );
        }
      }
      rowEl.appendChild(cellEl);
    }
    tableEl.appendChild(rowEl);
  }
  wrapperEl.appendChild(tableEl);
}

function applyCellStyle(
  cellEl: HTMLTableCellElement,
  cellProperties: Element | null,
  context: RenderContext,
): void {
  const fill = readFill(cellProperties, context.theme);
  const background = fill ? fillToCssBackground(fill) : null;
  if (background) {
    cellEl.style.background = background;
  }
  const margin = (name: string, fallback: number): string =>
    `${emuToPixels(numberAttribute(cellProperties, name, fallback), context)}px`;
  cellEl.style.padding = `${margin("marT", 45_720)} ${margin("marR", 91_440)} ${margin("marB", 45_720)} ${margin("marL", 91_440)}`;
  const anchor = attribute(cellProperties, "anchor");
  cellEl.style.verticalAlign =
    anchor === "ctr" ? "middle" : anchor === "b" ? "bottom" : "top";
  const borders: Array<[string, "borderLeft" | "borderRight" | "borderTop" | "borderBottom"]> = [
    ["lnL", "borderLeft"],
    ["lnR", "borderRight"],
    ["lnT", "borderTop"],
    ["lnB", "borderBottom"],
  ];
  for (const [name, property] of borders) {
    const line = firstChildNamed(cellProperties, name);
    if (!line) {
      continue;
    }
    const lineFill = readFill(line, context.theme);
    const color = lineFill?.type === "none" ? "transparent" : fillToCssColor(lineFill);
    if (color) {
      const width = Math.max(
        0.5,
        emuToPixels(numberAttribute(line, "w", 12_700), context),
      );
      cellEl.style[property] = `${width}px solid ${color}`;
    }
  }
}

async function yieldIfNeeded(context: RenderContext): Promise<void> {
  ensureRenderActive(context);
  const elapsed = context.now() - context.lastYieldAt;
  context.diagnostics.maxWorkSliceMs = Math.max(
    context.diagnostics.maxWorkSliceMs,
    elapsed,
  );
  if (elapsed < context.timeSliceMs) {
    return;
  }
  await context.yieldControl();
  context.diagnostics.yieldCount += 1;
  context.lastYieldAt = context.now();
  ensureRenderActive(context);
}

function recordWorkSlice(context: RenderContext): void {
  context.diagnostics.maxWorkSliceMs = Math.max(
    context.diagnostics.maxWorkSliceMs,
    context.now() - context.lastYieldAt,
  );
}

function createRenderYield(ownerDocument: Document): () => Promise<void> {
  return async () => {
    const renderWindow = ownerDocument.defaultView;
    if (!renderWindow) {
      await Promise.resolve();
      return;
    }
    await new Promise<void>((resolve) => {
      let completed = false;
      const finish = (): void => {
        if (completed) {
          return;
        }
        completed = true;
        renderWindow.clearTimeout(timeoutId);
        renderWindow.cancelAnimationFrame(frameId);
        resolve();
      };
      const timeoutId = renderWindow.setTimeout(finish, 50);
      const frameId = renderWindow.requestAnimationFrame(finish);
    });
  };
}

function renderUnsupportedPlaceholder(
  wrapperEl: HTMLElement,
  context: RenderContext,
  label: string,
): void {
  context.diagnostics.unsupportedObjectCount += 1;
  const placeholderEl = context.ownerDocument.createElement("div");
  placeholderEl.className = "pptx-reader-unsupported-object";
  placeholderEl.textContent = label;
  wrapperEl.appendChild(placeholderEl);
}

function hidesMasterShapes(document: Document | null): boolean {
  return attribute(document?.documentElement ?? null, "showMasterSp") === "0";
}

function isLineShape(shape: Element, inherited: Element[]): boolean {
  if (localName(shape) === "cxnSp") {
    return true;
  }
  const preset = [shape, ...inherited]
    .map((candidate) =>
      attribute(firstChildNamed(getShapeProperties(candidate), "prstGeom"), "prst"),
    )
    .find((value) => value !== null);
  return preset ? isLinePreset(preset) : false;
}

function getShapeProperties(shape: Element | null): Element | null {
  if (!shape) {
    return null;
  }
  switch (localName(shape)) {
    case "grpSp":
      return firstChildNamed(shape, "grpSpPr");
    default:
      return firstChildNamed(shape, "spPr");
  }
}

function parseShapeTransform(
  shape: Element,
  inherited: Element[],
  coordinateSpace: CoordinateSpace,
): ShapeTransform | null {
  const transform = [shape, ...inherited]
    .map(getTransform)
    .find((candidate) => candidate !== null);
  if (!transform) {
    return null;
  }
  const offset = firstChildNamed(transform, "off");
  const extent = firstChildNamed(transform, "ext");
  if (!offset || !extent) {
    return null;
  }
  const x =
    coordinateSpace.offsetX +
    numberAttribute(offset, "x") * coordinateSpace.scaleX;
  const y =
    coordinateSpace.offsetY +
    numberAttribute(offset, "y") * coordinateSpace.scaleY;
  return {
    x,
    y,
    width: numberAttribute(extent, "cx") * coordinateSpace.scaleX,
    height: numberAttribute(extent, "cy") * coordinateSpace.scaleY,
    rotation: numberAttribute(transform, "rot", 0) / 60_000,
    flipHorizontal: attribute(transform, "flipH") === "1",
    flipVertical: attribute(transform, "flipV") === "1",
  };
}

function parseGroupTransform(
  shape: Element,
  parentSpace: CoordinateSpace,
): CoordinateSpace | null {
  const transform = getTransform(shape);
  if (!transform) {
    return null;
  }
  const offset = firstChildNamed(transform, "off");
  const extent = firstChildNamed(transform, "ext");
  const childOffset = firstChildNamed(transform, "chOff");
  const childExtent = firstChildNamed(transform, "chExt");
  if (!offset || !extent || !childOffset || !childExtent) {
    return null;
  }
  const childWidth = numberAttribute(childExtent, "cx", 1);
  const childHeight = numberAttribute(childExtent, "cy", 1);
  const groupScaleX =
    (numberAttribute(extent, "cx") / Math.max(childWidth, 1)) *
    parentSpace.scaleX;
  const groupScaleY =
    (numberAttribute(extent, "cy") / Math.max(childHeight, 1)) *
    parentSpace.scaleY;
  return {
    offsetX:
      parentSpace.offsetX +
      numberAttribute(offset, "x") * parentSpace.scaleX -
      numberAttribute(childOffset, "x") * groupScaleX,
    offsetY:
      parentSpace.offsetY +
      numberAttribute(offset, "y") * parentSpace.scaleY -
      numberAttribute(childOffset, "y") * groupScaleY,
    scaleX: groupScaleX,
    scaleY: groupScaleY,
  };
}

function getTransform(shape: Element | null): Element | null {
  if (shape && localName(shape) === "graphicFrame") {
    return firstChildNamed(shape, "xfrm");
  }
  return firstChildNamed(getShapeProperties(shape), "xfrm");
}

function applyTransform(
  element: HTMLElement,
  transform: ShapeTransform,
  context: RenderContext,
): void {
  element.style.left = `${toPixelsX(transform.x, context)}px`;
  element.style.top = `${toPixelsY(transform.y, context)}px`;
  element.style.width = `${toPixelsX(transform.width, context)}px`;
  element.style.height = `${toPixelsY(transform.height, context)}px`;
  const transforms: string[] = [];
  if (transform.rotation !== 0) {
    transforms.push(`rotate(${transform.rotation}deg)`);
  }
  if (transform.flipHorizontal || transform.flipVertical) {
    transforms.push(
      `scale(${transform.flipHorizontal ? -1 : 1}, ${transform.flipVertical ? -1 : 1})`,
    );
  }
  if (transforms.length > 0) {
    element.style.transform = transforms.join(" ");
  }
}

function toPixelsX(value: number, context: RenderContext): number {
  return (value / context.slideWidth) * context.pixelWidth;
}

function toPixelsY(value: number, context: RenderContext): number {
  return (value / context.slideHeight) * context.pixelHeight;
}

function emuToPixels(value: number, context: RenderContext): number {
  return (value / context.slideWidth) * context.pixelWidth;
}

function pointsToPixels(points: number, context: RenderContext): number {
  return emuToPixels(points * EMU_PER_POINT, context);
}

function createBlob(bytes: Uint8Array, type: string): Blob {
  const buffer = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(buffer).set(bytes);
  return new Blob([buffer], { type });
}

function getGraphicFrameObjectType(shape: Element): string {
  if (firstDescendantNamed(shape, "chart")) {
    return "Chart";
  }
  const graphicData = firstDescendantNamed(shape, "graphicData");
  const uri = attribute(graphicData, "uri") ?? "";
  if (uri.includes("/diagram")) {
    return "SmartArt";
  }
  if (uri.endsWith("/presentationml/2006/ole")) {
    return "Embedded object";
  }
  return "Unsupported object";
}

function ensureRenderActive(context: RenderContext): void {
  if (context.isCancelled()) {
    throw new PptxRenderCancelledError();
  }
}
