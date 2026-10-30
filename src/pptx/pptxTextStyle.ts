import {
  type PptxTheme,
  fillToCssColor,
  readColorChoice,
  readFill,
  resolveTypeface,
} from "./pptxTheme";
import {
  attribute,
  childElements,
  firstChildNamed,
  firstDescendantNamed,
  localName,
  numberAttribute,
} from "./xml";

export interface PlaceholderIndex {
  byIndex: Map<string, Element>;
  byType: Map<string, Element>;
}

export type PptxSpacing =
  | { type: "percent"; value: number }
  | { type: "points"; value: number };

export interface PptxBulletStyle {
  kind: "char" | "number";
  character: string;
  numberScheme: string;
  startAt: number;
  typeface: string | null;
  color: string | null;
  sizePercent: number | null;
  sizePoints: number | null;
}

export interface PptxParagraphStyle {
  level: number;
  alignment: string;
  marginLeft: number;
  indent: number;
  lineSpacing: PptxSpacing | null;
  spaceBefore: PptxSpacing | null;
  spaceAfter: PptxSpacing | null;
  bullet: PptxBulletStyle | null;
  levelProperties: Element[];
}

export interface PptxRunStyle {
  sizePoints: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  capitals: "none" | "all" | "small";
  baseline: number;
  spacingPoints: number;
  color: string | null;
  highlight: string | null;
  typefaces: string[];
}

const DEFAULT_FONT_SIZE_POINTS = 18;
const TITLE_TYPES = new Set(["title", "ctrTitle"]);
const OTHER_TYPES = new Set(["dt", "ftr", "sldNum", "hdr"]);
const BULLET_TYPES = ["buNone", "buChar", "buAutoNum", "buBlip"];

// Symbol-font bullets store private-use code points that only render with the
// original font installed. Map the common ones to Unicode equivalents.
const SYMBOL_BULLETS: Record<string, Record<string, string>> = {
  Wingdings: {
    "§": "■",
    "ü": "✓",
    "Ø": "➢",
    "q": "❑",
    "n": "■",
    "l": "●",
    "v": "❖",
    "Ö": "➔",
    "à": "➔",
    "o": "□",
  },
  Symbol: {
    "·": "•",
    "-": "–",
  },
};

export function indexPlaceholders(document: Document | null): PlaceholderIndex {
  const result: PlaceholderIndex = { byIndex: new Map(), byType: new Map() };
  const shapeTree = firstDescendantNamed(document?.documentElement ?? null, "spTree");
  if (!shapeTree) {
    return result;
  }
  for (const shape of childElements(shapeTree)) {
    const placeholder = getPlaceholder(shape);
    if (!placeholder) {
      continue;
    }
    const index = attribute(placeholder, "idx");
    if (index && !result.byIndex.has(index)) {
      result.byIndex.set(index, shape);
    }
    const type = placeholderType(placeholder);
    if (!result.byType.has(type)) {
      result.byType.set(type, shape);
    }
  }
  return result;
}

/** Returns the `p:ph` element from a shape's non-visual properties. */
export function getPlaceholder(shape: Element | null): Element | null {
  if (!shape) {
    return null;
  }
  const nonVisual = childElements(shape).find((child) =>
    localName(child).startsWith("nv"),
  );
  return firstChildNamed(firstChildNamed(nonVisual, "nvPr"), "ph");
}

/**
 * Finds the layout and master shapes a slide placeholder inherits from, most
 * specific first. Layout placeholders are passed with an empty layout index.
 */
export function findInheritedPlaceholders(
  shape: Element,
  layout: PlaceholderIndex,
  master: PlaceholderIndex,
): Element[] {
  const placeholder = getPlaceholder(shape);
  if (!placeholder) {
    return [];
  }
  const index = attribute(placeholder, "idx");
  const type = placeholderType(placeholder);
  const layoutShape =
    (index ? layout.byIndex.get(index) : undefined) ??
    layout.byType.get(type) ??
    layout.byType.get(masterPlaceholderType(type));
  const masterType = masterPlaceholderType(
    placeholderType(getPlaceholder(layoutShape ?? null) ?? placeholder),
  );
  const masterShape =
    master.byType.get(masterType) ??
    (index ? master.byIndex.get(index) : undefined);
  return [layoutShape, masterShape].filter(
    (candidate): candidate is Element => candidate !== undefined && candidate !== shape,
  );
}

/** Selects the master `txStyles` entry used by a shape's text. */
export function masterTextStyle(
  shape: Element,
  master: Document | null,
): Element | null {
  const textStyles = firstDescendantNamed(master?.documentElement ?? null, "txStyles");
  const placeholder = getPlaceholder(shape);
  if (!placeholder) {
    return firstChildNamed(textStyles, "otherStyle");
  }
  const type = placeholderType(placeholder);
  if (TITLE_TYPES.has(type)) {
    return firstChildNamed(textStyles, "titleStyle");
  }
  if (OTHER_TYPES.has(type)) {
    return firstChildNamed(textStyles, "otherStyle");
  }
  return firstChildNamed(textStyles, "bodyStyle");
}

export function levelProperties(listStyles: Element[], level: number): Element[] {
  const name = `lvl${level + 1}pPr`;
  return listStyles
    .map((listStyle) => firstChildNamed(listStyle, name))
    .filter((properties): properties is Element => properties !== null);
}

export function resolveParagraphStyle(
  paragraph: Element,
  listStyles: Element[],
  theme: PptxTheme,
): PptxParagraphStyle {
  const ownProperties = firstChildNamed(paragraph, "pPr");
  const level = Math.min(8, Math.max(0, numberAttribute(ownProperties, "lvl", 0)));
  const sources = [
    ...(ownProperties ? [ownProperties] : []),
    ...levelProperties(listStyles, level),
  ];
  const first = (name: string): string | null => {
    for (const source of sources) {
      const value = attribute(source, name);
      if (value !== null) {
        return value;
      }
    }
    return null;
  };
  const child = (name: string): Element | null => {
    for (const source of sources) {
      const value = firstChildNamed(source, name);
      if (value) {
        return value;
      }
    }
    return null;
  };

  return {
    level,
    alignment: first("algn") ?? "l",
    marginLeft: Number(first("marL") ?? 0) || 0,
    indent: Number(first("indent") ?? 0) || 0,
    lineSpacing: readSpacing(child("lnSpc")),
    spaceBefore: readSpacing(child("spcBef")),
    spaceAfter: readSpacing(child("spcAft")),
    bullet: resolveBullet(sources, theme),
    levelProperties: sources,
  };
}

/**
 * Resolves character formatting. Explicit run properties win, then the shape
 * style font reference, then the inherited paragraph-level defaults.
 */
export function resolveRunStyle(
  runProperties: Element | null,
  paragraphStyle: PptxParagraphStyle,
  fontReference: Element | null,
  theme: PptxTheme,
): PptxRunStyle {
  const defaults = paragraphStyle.levelProperties
    .map((properties) => firstChildNamed(properties, "defRPr"))
    .filter((properties): properties is Element => properties !== null);
  const sources = [...(runProperties ? [runProperties] : []), ...defaults];
  const first = (name: string): string | null => {
    for (const source of sources) {
      const value = attribute(source, name);
      if (value !== null) {
        return value;
      }
    }
    return null;
  };

  let color: string | null = null;
  const hyperlink = firstChildNamed(runProperties, "hlinkClick");
  const explicitFill = readFill(runProperties, theme);
  if (explicitFill) {
    color = fillToCssColor(explicitFill);
  } else if (hyperlink) {
    color = theme.colors.hlink ?? null;
  } else {
    color = readColorChoice(fontReference, theme);
    for (const source of defaults) {
      if (color) {
        break;
      }
      color = fillToCssColor(readFill(source, theme));
    }
  }

  const typefaces: string[] = [];
  const addTypeface = (value: string | null): void => {
    const resolved = resolveTypeface(value, theme);
    if (resolved && !typefaces.includes(resolved)) {
      typefaces.push(resolved);
    }
  };
  const ownLatin = firstChildNamed(runProperties, "latin");
  const inheritedLatin = defaults
    .map((source) => firstChildNamed(source, "latin"))
    .find((latin) => latin !== null);
  if (ownLatin) {
    addTypeface(attribute(ownLatin, "typeface"));
  } else if (fontReference) {
    const fontSet =
      attribute(fontReference, "idx") === "major" ? theme.majorFont : theme.minorFont;
    addTypeface(fontSet.latin);
  } else if (inheritedLatin) {
    addTypeface(attribute(inheritedLatin, "typeface"));
  } else {
    addTypeface(theme.minorFont.latin);
  }
  const eastAsianSource = sources.find((source) => firstChildNamed(source, "ea"));
  addTypeface(
    eastAsianSource
      ? attribute(firstChildNamed(eastAsianSource, "ea"), "typeface")
      : theme.minorFont.eastAsian,
  );

  const highlight = readColorChoice(firstChildNamed(runProperties, "highlight"), theme);
  const underline = first("u");
  const strike = first("strike");
  const capitals = first("cap");
  return {
    sizePoints: Number(first("sz") ?? 0) / 100 || DEFAULT_FONT_SIZE_POINTS,
    bold: isTrue(first("b")),
    italic: isTrue(first("i")),
    underline: hyperlink !== null || (underline !== null && underline !== "none"),
    strike: strike !== null && strike !== "noStrike",
    capitals: capitals === "all" || capitals === "small" ? capitals : "none",
    baseline: Number(first("baseline") ?? 0) / 1000 || 0,
    spacingPoints: Number(first("spc") ?? 0) / 100 || 0,
    color,
    highlight,
    typefaces,
  };
}

export function formatAutoNumber(scheme: string, value: number): string {
  const lower = scheme.startsWith("alphaLc") || scheme.startsWith("romanLc");
  let label: string;
  if (scheme.startsWith("alpha")) {
    label = alphabetic(value);
  } else if (scheme.startsWith("roman")) {
    label = roman(value);
  } else {
    label = String(value);
  }
  if (lower) {
    label = label.toLowerCase();
  }
  if (scheme.endsWith("ParenBoth")) {
    return `(${label})`;
  }
  if (scheme.endsWith("ParenR")) {
    return `${label})`;
  }
  if (scheme.endsWith("Plain")) {
    return label;
  }
  return `${label}.`;
}

function resolveBullet(sources: Element[], theme: PptxTheme): PptxBulletStyle | null {
  const bulletSource = sources.find((source) =>
    BULLET_TYPES.some((name) => firstChildNamed(source, name)),
  );
  if (!bulletSource) {
    return null;
  }
  const character = firstChildNamed(bulletSource, "buChar");
  const autoNumber = firstChildNamed(bulletSource, "buAutoNum");
  if (!character && !autoNumber) {
    return null;
  }
  const child = (name: string): Element | null => {
    for (const source of sources) {
      const value = firstChildNamed(source, name);
      if (value) {
        return value;
      }
    }
    return null;
  };
  const font = child("buFontTx") ? null : child("buFont");
  const typeface = resolveTypeface(attribute(font, "typeface"), theme);
  const rawCharacter = attribute(character, "char") ?? "•";
  const sizePercent = child("buSzPct");
  const sizePoints = child("buSzPts");
  return {
    kind: autoNumber ? "number" : "char",
    character:
      (typeface ? SYMBOL_BULLETS[typeface]?.[rawCharacter] : undefined) ??
      rawCharacter,
    numberScheme: attribute(autoNumber, "type") ?? "arabicPeriod",
    startAt: numberAttribute(autoNumber, "startAt", 1),
    typeface: typeface && !SYMBOL_BULLETS[typeface] ? typeface : null,
    color: child("buClrTx") ? null : readColorChoice(child("buClr"), theme),
    sizePercent: sizePercent ? numberAttribute(sizePercent, "val", 100_000) / 100_000 : null,
    sizePoints: sizePoints ? numberAttribute(sizePoints, "val", 0) / 100 : null,
  };
}

function readSpacing(element: Element | null): PptxSpacing | null {
  const percent = firstChildNamed(element, "spcPct");
  if (percent) {
    return { type: "percent", value: numberAttribute(percent, "val", 100_000) / 100_000 };
  }
  const points = firstChildNamed(element, "spcPts");
  if (points) {
    return { type: "points", value: numberAttribute(points, "val", 0) / 100 };
  }
  return null;
}

function placeholderType(placeholder: Element): string {
  return attribute(placeholder, "type") ?? "obj";
}

function masterPlaceholderType(type: string): string {
  if (TITLE_TYPES.has(type)) {
    return "title";
  }
  if (OTHER_TYPES.has(type)) {
    return type;
  }
  return "body";
}

function isTrue(value: string | null): boolean {
  return value === "1" || value === "true";
}

function alphabetic(value: number): string {
  let remaining = Math.max(1, Math.floor(value));
  let result = "";
  while (remaining > 0) {
    remaining -= 1;
    result = String.fromCharCode(65 + (remaining % 26)) + result;
    remaining = Math.floor(remaining / 26);
  }
  return result;
}

function roman(value: number): string {
  const numerals: Array<[number, string]> = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ];
  let remaining = Math.max(1, Math.floor(value));
  let result = "";
  for (const [amount, numeral] of numerals) {
    while (remaining >= amount) {
      result += numeral;
      remaining -= amount;
    }
  }
  return result;
}
