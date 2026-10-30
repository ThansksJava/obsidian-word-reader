import {
  attribute,
  childElements,
  firstChildNamed,
  firstDescendantNamed,
  localName,
  numberAttribute,
} from "./xml";

export interface PptxFontSet {
  latin: string | null;
  eastAsian: string | null;
}

export interface PptxTheme {
  colors: Record<string, string>;
  majorFont: PptxFontSet;
  minorFont: PptxFontSet;
  fillStyles: Element[];
  lineStyles: Element[];
  backgroundFillStyles: Element[];
}

export interface PptxGradientStop {
  offset: number;
  color: string;
}

export type PptxFill =
  | { type: "none" }
  | { type: "solid"; color: string }
  | {
      type: "gradient";
      stops: PptxGradientStop[];
      angle: number;
      radial: boolean;
    }
  | { type: "image"; relationshipId: string };

export interface PptxShadow {
  offsetX: number;
  offsetY: number;
  blur: number;
  color: string;
}

const DEFAULT_SCHEME_COLORS: Record<string, string> = {
  dk1: "#000000",
  lt1: "#ffffff",
  dk2: "#44546a",
  lt2: "#e7e6e6",
  accent1: "#4472c4",
  accent2: "#ed7d31",
  accent3: "#a5a5a5",
  accent4: "#ffc000",
  accent5: "#5b9bd5",
  accent6: "#70ad47",
  hlink: "#0563c1",
  folHlink: "#954f72",
};

const DEFAULT_COLOR_MAP: Record<string, string> = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

const PRESET_COLORS: Record<string, string> = {
  black: "#000000",
  white: "#ffffff",
  red: "#ff0000",
  green: "#008000",
  blue: "#0000ff",
  yellow: "#ffff00",
  gray: "#808080",
  grey: "#808080",
  orange: "#ffa500",
  purple: "#800080",
  darkGray: "#a9a9a9",
  lightGray: "#d3d3d3",
  darkBlue: "#00008b",
  darkRed: "#8b0000",
  darkGreen: "#006400",
  cyan: "#00ffff",
  magenta: "#ff00ff",
};

const COLOR_ELEMENT_NAMES = new Set([
  "srgbClr",
  "schemeClr",
  "sysClr",
  "prstClr",
  "scrgbClr",
  "hslClr",
]);

const FILL_ELEMENT_NAMES = new Set([
  "noFill",
  "solidFill",
  "gradFill",
  "pattFill",
  "blipFill",
  "grpFill",
]);

/**
 * Builds the effective theme for one slide. Color mappings are applied in
 * order: the master `clrMap`, then any layout and slide `overrideClrMapping`.
 */
export function parsePptxTheme(
  themeDocument: Document | null,
  colorMappings: Array<Element | null>,
): PptxTheme {
  const scheme = { ...DEFAULT_SCHEME_COLORS };
  const themeElement = themeDocument?.documentElement ?? null;
  const colorScheme = firstDescendantNamed(themeElement, "clrScheme");
  if (colorScheme) {
    for (const entry of childElements(colorScheme)) {
      const value = readColor(entry, { colors: scheme });
      if (value) {
        scheme[localName(entry)] = value;
      }
    }
  }

  const colorMap = { ...DEFAULT_COLOR_MAP };
  for (const mapping of colorMappings) {
    if (!mapping) {
      continue;
    }
    for (const key of Object.keys(DEFAULT_COLOR_MAP)) {
      const target = attribute(mapping, key);
      if (target) {
        colorMap[key] = target;
      }
    }
  }
  const colors = { ...scheme };
  for (const [key, target] of Object.entries(colorMap)) {
    const value = scheme[target];
    if (value) {
      colors[key] = value;
    }
  }

  const fontScheme = firstDescendantNamed(themeElement, "fontScheme");
  const formatScheme = firstDescendantNamed(themeElement, "fmtScheme");
  return {
    colors,
    majorFont: readFontSet(firstChildNamed(fontScheme, "majorFont")),
    minorFont: readFontSet(firstChildNamed(fontScheme, "minorFont")),
    fillStyles: childrenOf(firstChildNamed(formatScheme, "fillStyleLst")),
    lineStyles: childrenOf(firstChildNamed(formatScheme, "lnStyleLst")),
    backgroundFillStyles: childrenOf(
      firstChildNamed(formatScheme, "bgFillStyleLst"),
    ),
  };
}

/** Returns the layout or slide color-map override, if it overrides anything. */
export function findColorMapOverride(document: Document | null): Element | null {
  const override = firstDescendantNamed(
    document?.documentElement ?? null,
    "clrMapOvr",
  );
  return firstChildNamed(override, "overrideClrMapping");
}

export function findMasterColorMap(document: Document | null): Element | null {
  return firstChildNamed(document?.documentElement ?? null, "clrMap");
}

/**
 * Resolves a theme font reference such as `+mn-lt` to a concrete typeface.
 */
export function resolveTypeface(
  typeface: string | null,
  theme: Pick<PptxTheme, "majorFont" | "minorFont">,
): string | null {
  if (!typeface) {
    return null;
  }
  switch (typeface) {
    case "+mj-lt":
      return theme.majorFont.latin;
    case "+mn-lt":
      return theme.minorFont.latin;
    case "+mj-ea":
      return theme.majorFont.eastAsian;
    case "+mn-ea":
      return theme.minorFont.eastAsian;
    default:
      return typeface.startsWith("+") ? null : typeface;
  }
}

/** Reads the first color child of `container`, applying its modifiers. */
export function readColorChoice(
  container: Element | null | undefined,
  theme: Pick<PptxTheme, "colors">,
  placeholderColor?: string,
): string | null {
  if (!container) {
    return null;
  }
  const colorElement = childElements(container).find((element) =>
    COLOR_ELEMENT_NAMES.has(localName(element)),
  );
  return colorElement ? readColor(colorElement, theme, placeholderColor) : null;
}

/**
 * Reads the fill child of a property container (`spPr`, `rPr`, `tcPr`, ...).
 * Returns null when the container does not specify a fill.
 */
export function readFill(
  container: Element | null | undefined,
  theme: Pick<PptxTheme, "colors">,
  placeholderColor?: string,
): PptxFill | null {
  if (!container) {
    return null;
  }
  const fillElement = childElements(container).find((element) =>
    FILL_ELEMENT_NAMES.has(localName(element)),
  );
  return fillElement ? readFillElement(fillElement, theme, placeholderColor) : null;
}

export function readFillElement(
  fillElement: Element,
  theme: Pick<PptxTheme, "colors">,
  placeholderColor?: string,
): PptxFill | null {
  switch (localName(fillElement)) {
    case "noFill":
      return { type: "none" };
    case "solidFill": {
      const color = readColorChoice(fillElement, theme, placeholderColor);
      return color ? { type: "solid", color } : null;
    }
    case "gradFill": {
      const stops = childrenOf(firstChildNamed(fillElement, "gsLst"))
        .map((stop) => ({
          offset: Math.min(1, Math.max(0, numberAttribute(stop, "pos") / 100_000)),
          color: readColorChoice(stop, theme, placeholderColor),
        }))
        .filter((stop): stop is PptxGradientStop => stop.color !== null)
        .sort((left, right) => left.offset - right.offset);
      if (stops.length === 0) {
        return null;
      }
      const linear = firstChildNamed(fillElement, "lin");
      return {
        type: "gradient",
        stops,
        angle: numberAttribute(linear, "ang", 0) / 60_000,
        radial: firstChildNamed(fillElement, "path") !== null,
      };
    }
    case "pattFill": {
      const foreground =
        readColorChoice(firstChildNamed(fillElement, "fgClr"), theme, placeholderColor) ??
        "#000000";
      const background =
        readColorChoice(firstChildNamed(fillElement, "bgClr"), theme, placeholderColor) ??
        "#ffffff";
      return {
        type: "solid",
        color: mixColors(foreground, background, patternDensity(attribute(fillElement, "prst"))),
      };
    }
    case "blipFill": {
      const blip = firstChildNamed(fillElement, "blip");
      const relationshipId = blip ? attribute(blip, "embed") : null;
      return relationshipId ? { type: "image", relationshipId } : null;
    }
    default:
      return null;
  }
}

/**
 * Resolves a style-matrix reference (`fillRef`, `bgRef`) against the theme.
 * Index 0 means "no fill"; indexes above 1000 select background fills.
 */
export function resolveFillReference(
  reference: Element | null,
  theme: PptxTheme,
): PptxFill | null {
  if (!reference) {
    return null;
  }
  const index = numberAttribute(reference, "idx", 0);
  if (index === 0) {
    return { type: "none" };
  }
  const placeholderColor = readColorChoice(reference, theme) ?? undefined;
  const style =
    index >= 1001
      ? theme.backgroundFillStyles[index - 1001]
      : theme.fillStyles[index - 1];
  const fill = style ? readFillElement(style, theme, placeholderColor) : null;
  if (fill) {
    return fill;
  }
  return placeholderColor ? { type: "solid", color: placeholderColor } : null;
}

export function fillToCssColor(fill: PptxFill | null): string | null {
  if (!fill) {
    return null;
  }
  switch (fill.type) {
    case "none":
      return "transparent";
    case "solid":
      return fill.color;
    case "gradient":
      return fill.stops[0]?.color ?? null;
    default:
      return null;
  }
}

export function fillToCssBackground(fill: PptxFill): string | null {
  switch (fill.type) {
    case "solid":
      return fill.color;
    case "gradient": {
      const stops = fill.stops
        .map((stop) => `${stop.color} ${Math.round(stop.offset * 100)}%`)
        .join(", ");
      return fill.radial
        ? `radial-gradient(circle, ${stops})`
        : `linear-gradient(${fill.angle + 90}deg, ${stops})`;
    }
    default:
      return null;
  }
}

export function readOuterShadow(
  effectList: Element | null,
  theme: Pick<PptxTheme, "colors">,
): PptxShadow | null {
  const shadow = firstChildNamed(effectList, "outerShdw");
  if (!shadow) {
    return null;
  }
  const color = readColorChoice(shadow, theme);
  if (!color) {
    return null;
  }
  const distance = numberAttribute(shadow, "dist", 0);
  const direction = (numberAttribute(shadow, "dir", 0) / 60_000) * (Math.PI / 180);
  return {
    offsetX: distance * Math.cos(direction),
    offsetY: distance * Math.sin(direction),
    blur: numberAttribute(shadow, "blurRad", 0),
    color,
  };
}

function readColor(
  colorElement: Element,
  theme: Pick<PptxTheme, "colors">,
  placeholderColor?: string,
): string | null {
  let color: string | undefined;
  switch (localName(colorElement)) {
    case "srgbClr":
      color = normalizeHexColor(attribute(colorElement, "val"));
      break;
    case "schemeClr": {
      const name = attribute(colorElement, "val") ?? "";
      color = name === "phClr" ? placeholderColor : theme.colors[name];
      break;
    }
    case "sysClr":
      color = normalizeHexColor(
        attribute(colorElement, "lastClr") ?? attribute(colorElement, "val"),
      );
      break;
    case "prstClr":
      color = PRESET_COLORS[attribute(colorElement, "val") ?? ""];
      break;
    case "scrgbClr":
      color = rgbToHex(
        (numberAttribute(colorElement, "r") / 100_000) * 255,
        (numberAttribute(colorElement, "g") / 100_000) * 255,
        (numberAttribute(colorElement, "b") / 100_000) * 255,
      );
      break;
    case "hslClr": {
      const [red, green, blue] = hslToRgb(
        numberAttribute(colorElement, "hue") / 60_000 / 360,
        numberAttribute(colorElement, "sat") / 100_000,
        numberAttribute(colorElement, "lum") / 100_000,
      );
      color = rgbToHex(red, green, blue);
      break;
    }
    default:
      // A nested scheme entry such as <a:dk1><a:srgbClr/></a:dk1>.
      return readColorChoice(colorElement, theme, placeholderColor);
  }
  if (!color) {
    return null;
  }
  // Placeholder colors may already carry alpha from an outer reference.
  if (color.startsWith("rgba(")) {
    return color;
  }
  return applyColorModifiers(color, colorElement);
}

function applyColorModifiers(color: string, colorElement: Element): string {
  let [red, green, blue] = hexToRgb(color);
  let alpha = 1;
  for (const modifier of childElements(colorElement)) {
    const amount = numberAttribute(modifier, "val", 0) / 100_000;
    switch (localName(modifier)) {
      case "tint":
        red += (255 - red) * (1 - amount);
        green += (255 - green) * (1 - amount);
        blue += (255 - blue) * (1 - amount);
        break;
      case "shade":
        red *= amount;
        green *= amount;
        blue *= amount;
        break;
      case "lumMod":
      case "lumOff":
      case "satMod": {
        const [hue, saturation, luminance] = rgbToHsl(red, green, blue);
        const next: [number, number, number] =
          localName(modifier) === "lumMod"
            ? [hue, saturation, luminance * amount]
            : localName(modifier) === "lumOff"
              ? [hue, saturation, luminance + amount]
              : [hue, saturation * amount, luminance];
        [red, green, blue] = hslToRgb(
          next[0],
          Math.min(1, Math.max(0, next[1])),
          Math.min(1, Math.max(0, next[2])),
        );
        break;
      }
      case "alpha":
        alpha = Math.min(1, Math.max(0, amount));
        break;
      case "alphaMod":
        alpha = Math.min(1, Math.max(0, alpha * amount));
        break;
      default:
        break;
    }
  }
  const hex = rgbToHex(red, green, blue);
  if (alpha >= 1) {
    return hex;
  }
  const [r, g, b] = hexToRgb(hex);
  return `rgba(${r}, ${g}, ${b}, ${Math.round(alpha * 1000) / 1000})`;
}

function readFontSet(font: Element | null): PptxFontSet {
  const latin = attribute(firstChildNamed(font, "latin"), "typeface");
  const eastAsian = attribute(firstChildNamed(font, "ea"), "typeface");
  return {
    latin: latin || null,
    eastAsian: eastAsian || null,
  };
}

function patternDensity(preset: string | null): number {
  const match = /^pct(\d+)$/.exec(preset ?? "");
  return match ? Number(match[1]) / 100 : 0.5;
}

function mixColors(foreground: string, background: string, amount: number): string {
  if (foreground.startsWith("rgba(") || background.startsWith("rgba(")) {
    return foreground;
  }
  const [fr, fg, fb] = hexToRgb(foreground);
  const [br, bg, bb] = hexToRgb(background);
  return rgbToHex(
    br + (fr - br) * amount,
    bg + (fg - bg) * amount,
    bb + (fb - bb) * amount,
  );
}

function childrenOf(element: Element | null): Element[] {
  return element ? childElements(element) : [];
}

function normalizeHexColor(value: string | null): string | undefined {
  if (!value || !/^[0-9a-f]{6}$/i.test(value)) {
    return undefined;
  }
  return `#${value.toLowerCase()}`;
}

function hexToRgb(color: string): [number, number, number] {
  return [
    Number.parseInt(color.slice(1, 3), 16),
    Number.parseInt(color.slice(3, 5), 16),
    Number.parseInt(color.slice(5, 7), 16),
  ];
}

function rgbToHex(red: number, green: number, blue: number): string {
  return `#${[red, green, blue]
    .map((value) =>
      Math.min(255, Math.max(0, Math.round(value)))
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")}`;
}

function rgbToHsl(red: number, green: number, blue: number): [number, number, number] {
  const r = red / 255;
  const g = green / 255;
  const b = blue / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const luminance = (max + min) / 2;
  if (max === min) {
    return [0, 0, luminance];
  }
  const delta = max - min;
  const saturation =
    luminance > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue: number;
  if (max === r) {
    hue = (g - b) / delta + (g < b ? 6 : 0);
  } else if (max === g) {
    hue = (b - r) / delta + 2;
  } else {
    hue = (r - g) / delta + 4;
  }
  return [hue / 6, saturation, luminance];
}

function hslToRgb(hue: number, saturation: number, luminance: number): [number, number, number] {
  if (saturation === 0) {
    const gray = luminance * 255;
    return [gray, gray, gray];
  }
  const q =
    luminance < 0.5
      ? luminance * (1 + saturation)
      : luminance + saturation - luminance * saturation;
  const p = 2 * luminance - q;
  const channel = (offset: number): number => {
    let t = hue + offset;
    if (t < 0) {
      t += 1;
    }
    if (t > 1) {
      t -= 1;
    }
    if (t < 1 / 6) {
      return p + (q - p) * 6 * t;
    }
    if (t < 1 / 2) {
      return q;
    }
    if (t < 2 / 3) {
      return p + (q - p) * (2 / 3 - t) * 6;
    }
    return p;
  };
  return [channel(1 / 3) * 255, channel(0) * 255, channel(-1 / 3) * 255];
}
