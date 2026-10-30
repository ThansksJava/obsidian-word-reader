import {
  attribute,
  childElements,
  childrenNamed,
  firstChildNamed,
  localName,
  numberAttribute,
} from "./xml";

export interface PptxPathSpec {
  d: string;
  fill: boolean;
  stroke: boolean;
}

/** Preset geometries that are open lines and therefore never filled. */
const LINE_PRESETS = new Set([
  "line",
  "straightConnector1",
  "bentConnector2",
  "bentConnector3",
  "bentConnector4",
  "bentConnector5",
  "curvedConnector2",
  "curvedConnector3",
  "curvedConnector4",
  "curvedConnector5",
  "arc",
  "leftBracket",
  "rightBracket",
  "leftBrace",
  "rightBrace",
]);

export function isLinePreset(preset: string): boolean {
  return LINE_PRESETS.has(preset);
}

/**
 * Builds SVG paths for a shape's geometry in pixel coordinates.
 * `shapeWidthEmu` and `shapeHeightEmu` define the custom-geometry path space
 * when a path does not declare its own size.
 */
export function buildGeometryPaths(
  shapeProperties: Element | null,
  width: number,
  height: number,
  shapeWidthEmu: number,
  shapeHeightEmu: number,
): PptxPathSpec[] {
  const customGeometry = firstChildNamed(shapeProperties, "custGeom");
  if (customGeometry) {
    const paths = buildCustomPaths(
      customGeometry,
      width,
      height,
      shapeWidthEmu,
      shapeHeightEmu,
    );
    if (paths.length > 0) {
      return paths;
    }
  }
  const presetGeometry = firstChildNamed(shapeProperties, "prstGeom");
  return buildPresetPaths(
    attribute(presetGeometry, "prst") ?? "rect",
    readAdjustments(firstChildNamed(presetGeometry, "avLst")),
    width,
    height,
  );
}

export function buildPresetPaths(
  preset: string,
  adjustments: Map<string, number>,
  w: number,
  h: number,
): PptxPathSpec[] {
  const ss = Math.min(w, h);
  const adj = (name: string, fallback: number): number =>
    (adjustments.get(name) ?? fallback) / 100_000;
  const filled = (d: string): PptxPathSpec[] => [{ d, fill: true, stroke: true }];
  const open = (d: string): PptxPathSpec[] => [{ d, fill: false, stroke: true }];

  switch (preset) {
    case "line":
    case "straightConnector1":
      return open(`M0,0 L${w},${h}`);
    case "bentConnector2":
      return open(`M0,0 L${w},0 L${w},${h}`);
    case "bentConnector3": {
      const x = w * adj("adj1", 50_000);
      return open(`M0,0 L${x},0 L${x},${h} L${w},${h}`);
    }
    case "curvedConnector3": {
      const x = w * adj("adj1", 50_000);
      return open(`M0,0 C${x},0 ${x},${h} ${w},${h}`);
    }
    case "ellipse":
    case "flowChartConnector":
      return filled(ellipsePath(w / 2, h / 2, w / 2, h / 2));
    case "roundRect":
    case "flowChartAlternateProcess": {
      const radius = ss * (preset === "roundRect" ? adj("adj", 16_667) : 0.16667);
      return filled(roundRectPath(w, h, radius, radius));
    }
    case "flowChartTerminator":
      return filled(roundRectPath(w, h, (w * 3475) / 21_600, h / 2));
    case "snip1Rect": {
      const s = ss * adj("adj", 16_667);
      return filled(`M0,0 L${w - s},0 L${w},${s} L${w},${h} L0,${h} Z`);
    }
    case "triangle": {
      const x = w * adj("adj", 50_000);
      return filled(`M${x},0 L${w},${h} L0,${h} Z`);
    }
    case "rtTriangle":
      return filled(`M0,0 L${w},${h} L0,${h} Z`);
    case "diamond":
    case "flowChartDecision":
      return filled(`M${w / 2},0 L${w},${h / 2} L${w / 2},${h} L0,${h / 2} Z`);
    case "parallelogram":
    case "flowChartInputOutput": {
      const o =
        preset === "parallelogram" ? ss * adj("adj", 25_000) : w / 5;
      return filled(`M${o},0 L${w},0 L${w - o},${h} L0,${h} Z`);
    }
    case "trapezoid": {
      const o = ss * adj("adj", 25_000);
      return filled(`M0,${h} L${o},0 L${w - o},0 L${w},${h} Z`);
    }
    case "pentagon":
      return filled(
        `M${w / 2},0 L${w},${h * 0.38} L${w * 0.81},${h} L${w * 0.19},${h} L0,${h * 0.38} Z`,
      );
    case "hexagon": {
      const o = ss * adj("adj", 25_000);
      return filled(
        `M0,${h / 2} L${o},0 L${w - o},0 L${w},${h / 2} L${w - o},${h} L${o},${h} Z`,
      );
    }
    case "octagon": {
      const o = ss * adj("adj", 29_289);
      return filled(
        `M${o},0 L${w - o},0 L${w},${o} L${w},${h - o} L${w - o},${h} L${o},${h} L0,${h - o} L0,${o} Z`,
      );
    }
    case "homePlate": {
      const o = ss * adj("adj", 50_000);
      return filled(`M0,0 L${w - o},0 L${w},${h / 2} L${w - o},${h} L0,${h} Z`);
    }
    case "chevron": {
      const o = ss * adj("adj", 50_000);
      return filled(
        `M0,0 L${w - o},0 L${w},${h / 2} L${w - o},${h} L0,${h} L${o},${h / 2} Z`,
      );
    }
    case "rightArrow":
    case "leftArrow": {
      const half = (h * adj("adj1", 50_000)) / 2;
      const head = ss * adj("adj2", 50_000);
      const y1 = h / 2 - half;
      const y2 = h / 2 + half;
      return preset === "rightArrow"
        ? filled(
            `M0,${y1} L${w - head},${y1} L${w - head},0 L${w},${h / 2} L${w - head},${h} L${w - head},${y2} L0,${y2} Z`,
          )
        : filled(
            `M${w},${y1} L${head},${y1} L${head},0 L0,${h / 2} L${head},${h} L${head},${y2} L${w},${y2} Z`,
          );
    }
    case "upArrow":
    case "downArrow": {
      const half = (w * adj("adj1", 50_000)) / 2;
      const head = ss * adj("adj2", 50_000);
      const x1 = w / 2 - half;
      const x2 = w / 2 + half;
      return preset === "downArrow"
        ? filled(
            `M${x1},0 L${x1},${h - head} L0,${h - head} L${w / 2},${h} L${w},${h - head} L${x2},${h - head} L${x2},0 Z`,
          )
        : filled(
            `M${x1},${h} L${x1},${head} L0,${head} L${w / 2},0 L${w},${head} L${x2},${head} L${x2},${h} Z`,
          );
    }
    case "leftRightArrow": {
      const half = (h * adj("adj1", 50_000)) / 2;
      const head = ss * adj("adj2", 50_000);
      const y1 = h / 2 - half;
      const y2 = h / 2 + half;
      return filled(
        `M0,${h / 2} L${head},0 L${head},${y1} L${w - head},${y1} L${w - head},0 L${w},${h / 2} L${w - head},${h} L${w - head},${y2} L${head},${y2} L${head},${h} Z`,
      );
    }
    case "plus": {
      const o = ss * adj("adj", 25_000);
      return filled(
        `M${o},0 L${w - o},0 L${w - o},${o} L${w},${o} L${w},${h - o} L${w - o},${h - o} L${w - o},${h} L${o},${h} L${o},${h - o} L0,${h - o} L0,${o} L${o},${o} Z`,
      );
    }
    case "star5":
      return filled(starPath(w, h, 5, 0.382));
    case "lightningBolt": {
      const points = [
        [8_472, 0], [12_860, 6_080], [11_050, 6_797], [16_577, 12_007],
        [14_767, 12_877], [21_600, 21_600], [10_012, 14_915], [12_222, 13_987],
        [5_022, 9_705], [7_602, 8_382], [0, 3_890],
      ];
      return filled(
        `${points
          .map(([x = 0, y = 0], index) =>
            `${index === 0 ? "M" : "L"}${(x * w) / 21_600},${(y * h) / 21_600}`)
          .join(" ")} Z`,
      );
    }
    case "can":
    case "flowChartMagneticDisk": {
      const radiusY =
        preset === "can" ? (ss * adj("adj", 25_000)) / 2 : h / 6;
      return cylinderPaths(w, h, radiusY);
    }
    case "flowChartMagneticTape":
      return [
        { d: ellipsePath(w / 2, h / 2, w / 2, h / 2), fill: true, stroke: true },
        { d: `M${w / 2},${h} L${w},${h} L${w},${h * 0.85}`, fill: false, stroke: true },
      ];
    case "flowChartDocument": {
      const sx = w / 21_600;
      const sy = h / 21_600;
      return filled(
        `M0,0 L${w},0 L${w},${17_322 * sy} C${10_800 * sx},${17_322 * sy} ${10_800 * sx},${23_922 * sy} 0,${20_172 * sy} Z`,
      );
    }
    case "flowChartDisplay":
      return filled(
        `M0,${h / 2} L${w / 6},0 L${(w * 5) / 6},0 A${w / 6},${h / 2} 0 0 1 ${(w * 5) / 6},${h} L${w / 6},${h} Z`,
      );
    case "flowChartPredefinedProcess":
      return [
        { d: rectPath(w, h), fill: true, stroke: true },
        { d: `M${w / 8},0 L${w / 8},${h} M${(w * 7) / 8},0 L${(w * 7) / 8},${h}`, fill: false, stroke: true },
      ];
    case "cube": {
      const o = ss * adj("adj", 25_000);
      return [
        { d: `M0,${o} L${o},0 L${w},0 L${w},${h - o} L${w - o},${h} L0,${h} Z`, fill: true, stroke: true },
        { d: `M0,${o} L${w - o},${o} L${w},0 M${w - o},${o} L${w - o},${h}`, fill: false, stroke: true },
      ];
    }
    case "leftBracket":
      return open(`M${w},0 Q0,0 0,${h * 0.1} L0,${h * 0.9} Q0,${h} ${w},${h}`);
    case "rightBracket":
      return open(`M0,0 Q${w},0 ${w},${h * 0.1} L${w},${h * 0.9} Q${w},${h} 0,${h}`);
    case "leftBrace":
    case "rightBrace": {
      const curl = Math.min(h * adj("adj1", 8_333), h / 4);
      const middle = h * adj("adj2", 50_000);
      const edge = preset === "rightBrace" ? 0 : w;
      const tip = preset === "rightBrace" ? w : 0;
      const spine = w / 2;
      return open(
        `M${edge},0 Q${spine},0 ${spine},${curl} L${spine},${middle - curl} Q${spine},${middle} ${tip},${middle} Q${spine},${middle} ${spine},${middle + curl} L${spine},${h - curl} Q${spine},${h} ${edge},${h}`,
      );
    }
    case "rect":
    case "flowChartProcess":
    default:
      return filled(rectPath(w, h));
  }
}

function buildCustomPaths(
  customGeometry: Element,
  width: number,
  height: number,
  shapeWidthEmu: number,
  shapeHeightEmu: number,
): PptxPathSpec[] {
  const guides = readAdjustments(firstChildNamed(customGeometry, "avLst"));
  for (const [name, value] of readAdjustments(
    firstChildNamed(customGeometry, "gdLst"),
  )) {
    guides.set(name, value);
  }
  guides.set("w", shapeWidthEmu);
  guides.set("h", shapeHeightEmu);
  const pathList = firstChildNamed(customGeometry, "pathLst");
  if (!pathList) {
    return [];
  }
  const result: PptxPathSpec[] = [];
  for (const path of childrenNamed(pathList, "path")) {
    const pathWidth = numberAttribute(path, "w", 0) || shapeWidthEmu || 1;
    const pathHeight = numberAttribute(path, "h", 0) || shapeHeightEmu || 1;
    const scaleX = width / pathWidth;
    const scaleY = height / pathHeight;
    const value = (element: Element, name: string): number => {
      const raw = attribute(element, name) ?? "0";
      const numeric = Number(raw);
      return Number.isFinite(numeric) ? numeric : guides.get(raw) ?? 0;
    };
    const commands: string[] = [];
    let currentX = 0;
    let currentY = 0;
    for (const command of childElements(path)) {
      const points = childrenNamed(command, "pt").map(
        (point) => [value(point, "x"), value(point, "y")] as const,
      );
      switch (localName(command)) {
        case "moveTo":
        case "lnTo": {
          const point = points[0];
          if (!point) {
            break;
          }
          [currentX, currentY] = point;
          commands.push(
            `${localName(command) === "moveTo" ? "M" : "L"}${currentX * scaleX},${currentY * scaleY}`,
          );
          break;
        }
        case "cubicBezTo":
        case "quadBezTo": {
          const last = points.at(-1);
          if (!last || points.length < (localName(command) === "cubicBezTo" ? 3 : 2)) {
            break;
          }
          commands.push(
            `${localName(command) === "cubicBezTo" ? "C" : "Q"}${points
              .map(([x, y]) => `${x * scaleX},${y * scaleY}`)
              .join(" ")}`,
          );
          [currentX, currentY] = last;
          break;
        }
        case "arcTo": {
          const arc = arcCommand(
            currentX,
            currentY,
            value(command, "wR"),
            value(command, "hR"),
            value(command, "stAng") / 60_000,
            value(command, "swAng") / 60_000,
            scaleX,
            scaleY,
          );
          commands.push(arc.command);
          currentX = arc.endX;
          currentY = arc.endY;
          break;
        }
        case "close":
          commands.push("Z");
          break;
        default:
          break;
      }
    }
    if (commands.length === 0) {
      continue;
    }
    const strokeValue = attribute(path, "stroke");
    result.push({
      d: commands.join(" "),
      fill: attribute(path, "fill") !== "none",
      stroke: strokeValue !== "0" && strokeValue !== "false",
    });
  }
  return result;
}

function arcCommand(
  startX: number,
  startY: number,
  radiusX: number,
  radiusY: number,
  startAngle: number,
  sweepAngle: number,
  scaleX: number,
  scaleY: number,
): { command: string; endX: number; endY: number } {
  if (radiusX <= 0 || radiusY <= 0 || sweepAngle === 0) {
    return { command: "", endX: startX, endY: startY };
  }
  const parametric = (angle: number): number => {
    const radians = (angle * Math.PI) / 180;
    return Math.atan2(radiusX * Math.sin(radians), radiusY * Math.cos(radians));
  };
  const start = parametric(startAngle);
  const centerX = startX - radiusX * Math.cos(start);
  const centerY = startY - radiusY * Math.sin(start);
  const pointAt = (angle: number): [number, number] => {
    const theta = parametric(angle);
    return [
      centerX + radiusX * Math.cos(theta),
      centerY + radiusY * Math.sin(theta),
    ];
  };
  const sweepFlag = sweepAngle > 0 ? 1 : 0;
  const segments: string[] = [];
  // SVG cannot draw a full ellipse with one arc, so split large sweeps.
  const steps = Math.abs(sweepAngle) > 180 ? 2 : 1;
  let endX = startX;
  let endY = startY;
  for (let step = 1; step <= steps; step += 1) {
    [endX, endY] = pointAt(startAngle + (sweepAngle * step) / steps);
    segments.push(
      `A${radiusX * scaleX},${radiusY * scaleY} 0 0 ${sweepFlag} ${endX * scaleX},${endY * scaleY}`,
    );
  }
  return { command: segments.join(" "), endX, endY };
}

function readAdjustments(list: Element | null): Map<string, number> {
  const result = new Map<string, number>();
  if (!list) {
    return result;
  }
  for (const guide of childrenNamed(list, "gd")) {
    const name = attribute(guide, "name");
    const match = /^val\s+(-?\d+(?:\.\d+)?)$/.exec(attribute(guide, "fmla") ?? "");
    if (name && match) {
      result.set(name, Number(match[1]));
    }
  }
  return result;
}

function rectPath(w: number, h: number): string {
  return `M0,0 L${w},0 L${w},${h} L0,${h} Z`;
}

function roundRectPath(w: number, h: number, radiusX: number, radiusY: number): string {
  const rx = Math.min(radiusX, w / 2);
  const ry = Math.min(radiusY, h / 2);
  return `M${rx},0 L${w - rx},0 A${rx},${ry} 0 0 1 ${w},${ry} L${w},${h - ry} A${rx},${ry} 0 0 1 ${w - rx},${h} L${rx},${h} A${rx},${ry} 0 0 1 0,${h - ry} L0,${ry} A${rx},${ry} 0 0 1 ${rx},0 Z`;
}

function ellipsePath(cx: number, cy: number, rx: number, ry: number): string {
  return `M${cx - rx},${cy} A${rx},${ry} 0 1 1 ${cx + rx},${cy} A${rx},${ry} 0 1 1 ${cx - rx},${cy} Z`;
}

function cylinderPaths(w: number, h: number, radiusY: number): PptxPathSpec[] {
  const rx = w / 2;
  const ry = Math.min(radiusY, h / 2);
  return [
    {
      d: `M0,${ry} A${rx},${ry} 0 0 1 ${w},${ry} L${w},${h - ry} A${rx},${ry} 0 0 1 0,${h - ry} Z`,
      fill: true,
      stroke: true,
    },
    {
      d: `M${w},${ry} A${rx},${ry} 0 0 1 0,${ry}`,
      fill: false,
      stroke: true,
    },
  ];
}

function starPath(w: number, h: number, points: number, innerRatio: number): string {
  const commands: string[] = [];
  for (let index = 0; index < points * 2; index += 1) {
    const radius = index % 2 === 0 ? 1 : innerRatio;
    const angle = -Math.PI / 2 + (index * Math.PI) / points;
    const x = w / 2 + (w / 2) * radius * Math.cos(angle);
    const y = h / 2 + (h / 2) * radius * Math.sin(angle);
    commands.push(`${index === 0 ? "M" : "L"}${x},${y}`);
  }
  return `${commands.join(" ")} Z`;
}
