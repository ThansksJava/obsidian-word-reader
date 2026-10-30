import assert from "node:assert/strict";
import test from "node:test";

import { extractMetafileBitmap } from "../src/pptx/pptxMetafile";
import { PptxPackage } from "../src/pptx/pptxPackage";
import { renderPptxSlide } from "../src/pptx/pptxRenderer";
import { numberAttribute, parseXml } from "../src/pptx/xml";
import { createStyleInheritancePptx } from "./pptxFixture";

class MockClassList {
  readonly values = new Set<string>();

  add(...names: string[]): void {
    for (const name of names) {
      this.values.add(name);
    }
  }
}

class MockElement {
  readonly children: MockElement[] = [];
  readonly attributes = new Map<string, string>();
  readonly classList = new MockClassList();
  readonly style: Record<string, string> = {};
  className = "";
  textContent = "";
  src = "";
  alt = "";
  loading = "";
  decoding = "";
  colSpan = 1;
  rowSpan = 1;

  constructor(readonly tagName: string) {}

  appendChild(child: MockElement): MockElement {
    this.children.push(child);
    return child;
  }

  hasChildNodes(): boolean {
    return this.children.length > 0;
  }

  setAttribute(name: string, value: string): void {
    this.attributes.set(name, value);
  }
}

class MockDocument {
  createElement(name: string): MockElement {
    return new MockElement(name);
  }

  createElementNS(_namespace: string, name: string): MockElement {
    return new MockElement(name);
  }
}

async function renderStyleFixture(): Promise<MockElement[]> {
  const pptx = await PptxPackage.load(await createStyleInheritancePptx());
  const rendered = await renderPptxSlide(
    pptx,
    await pptx.getSlideContext(0),
    new MockDocument() as unknown as Document,
  );
  for (const resource of rendered.resources) {
    URL.revokeObjectURL(resource);
  }
  return flatten(rendered.element as unknown as MockElement);
}

void test("renderPptxSlide applies layout color-map overrides and hides master shapes", async () => {
  const [stage, ...elements] = await renderStyleFixture();

  assert.equal(stage?.style.background, "#000000");
  assert.equal(stage?.style.color, "#ffffff");
  assert.ok(
    !elements.some((element) => element.attributes.get("fill") === "#cc5500"),
    "showMasterSp=0 on the layout hides the master decoration",
  );
});

void test("renderPptxSlide inherits placeholder position, bullets, and spacing", async () => {
  const elements = await renderStyleFixture();
  const body = elements.find(
    (element) =>
      element.className.includes("pptx-reader-sp") &&
      flatten(element).some((child) => child.textContent === "First level"),
  );
  assert.ok(body);
  assert.equal(body.style.left, "36px");
  assert.equal(body.style.top, "126px");

  const paragraphs = flatten(body).filter(
    (element) => element.className === "pptx-reader-paragraph",
  );
  const [first, second] = paragraphs;
  assert.ok(first && second);
  assert.equal(first.style.fontSize, "28px");
  assert.equal(first.style.paddingLeft, "18px");
  assert.equal(first.style.textIndent, "-18px");
  assert.equal(first.style.marginTop, undefined);
  assert.equal(second.style.fontSize, "24px");
  assert.equal(second.style.paddingLeft, "54px");
  assert.equal(second.style.marginTop, "10px");

  const bullets = paragraphs.map(
    (paragraph) =>
      paragraph.children.find((child) => child.className === "pptx-reader-bullet"),
  );
  assert.equal(bullets[0]?.textContent, "•");
  assert.equal(bullets[0]?.style.width, "18px");
  assert.equal(bullets[1]?.textContent, "–");

  const firstRun = first.children.find((child) => child.textContent === "First level");
  assert.equal(firstRun?.style.color, "#ffffff");
});

void test("renderPptxSlide applies autofit scaling and slide-number fields", async () => {
  const elements = await renderStyleFixture();
  const shrunk = elements.find((element) => element.textContent === "Shrunk");
  assert.equal(shrunk?.style.fontSize, "20px");
  assert.ok(elements.some((element) => element.textContent === "1"));
});

void test("renderPptxSlide draws zero-height connectors with dashes and arrowheads", async () => {
  const elements = await renderStyleFixture();
  const connector = elements.find((element) =>
    element.className.includes("pptx-reader-cxnSp"),
  );
  assert.ok(connector);
  assert.equal(connector.style.height, "0px");
  const path = flatten(connector).find((element) => element.tagName === "path" && element.attributes.has("marker-end"));
  assert.ok(path);
  assert.equal(path.attributes.get("stroke"), "#ff0000");
  assert.equal(path.attributes.get("stroke-width"), "2");
  assert.equal(path.attributes.get("stroke-dasharray"), "8 6");
  assert.ok(flatten(connector).some((element) => element.tagName === "marker"));
});

void test("renderPptxSlide draws custom geometry with gradient fills", async () => {
  const elements = await renderStyleFixture();
  const gradient = elements.find((element) => element.tagName === "linearGradient");
  assert.ok(gradient);
  const stops = gradient.children.map((stop) => stop.attributes.get("stop-color"));
  assert.deepEqual(stops, ["#00ff00", "rgba(0, 0, 255, 0.5)"]);
  const path = elements.find(
    (element) =>
      element.tagName === "path" &&
      element.attributes.get("fill") === `url(#${gradient.attributes.get("id")})`,
  );
  assert.equal(path?.attributes.get("d"), "M0,100 L50,0 C60,20 80,60 100,100 Z");
});

void test("numberAttribute uses the fallback for missing attributes", () => {
  const element = parseXml('<a x="5" empty=""/>', "test.xml").documentElement;
  assert.equal(numberAttribute(element, "x", 1), 5);
  assert.equal(numberAttribute(element, "missing", 7), 7);
  assert.equal(numberAttribute(element, "empty", 3), 3);
});

void test("extractMetafileBitmap converts bitmap-only EMF and WMF files to BMP", () => {
  const dib = createDib();
  for (const metafile of [createBitmapEmf(dib), createBitmapWmf(dib)]) {
    const bitmap = extractMetafileBitmap(metafile.bytes, metafile.extension);
    assert.ok(bitmap, metafile.extension);
    assert.equal(bitmap.mimeType, "image/bmp");
    const view = new DataView(bitmap.bytes.buffer);
    assert.equal(String.fromCharCode(bitmap.bytes[0] ?? 0, bitmap.bytes[1] ?? 0), "BM");
    assert.equal(view.getUint32(2, true), 14 + dib.byteLength);
    assert.equal(view.getUint32(10, true), 14 + 40);
    assert.deepEqual([...bitmap.bytes.subarray(14)], [...dib]);
  }
});

void test("extractMetafileBitmap rejects metafiles with vector drawing records", () => {
  const records = [emfRecord(1, new Uint8Array(80)), emfRecord(86, new Uint8Array(24)), emfRecord(14, new Uint8Array(12))];
  assert.equal(extractMetafileBitmap(concat(records), "emf"), null);
  assert.equal(extractMetafileBitmap(Uint8Array.from([1, 2, 3]), "emf"), null);
});

function createDib(): Uint8Array {
  const dib = new Uint8Array(40 + 4);
  const view = new DataView(dib.buffer);
  view.setUint32(0, 40, true);
  view.setInt32(4, 1, true);
  view.setInt32(8, 1, true);
  view.setUint16(12, 1, true);
  view.setUint16(14, 32, true);
  dib.set([0x10, 0x20, 0x30, 0x00], 40);
  return dib;
}

function createBitmapEmf(dib: Uint8Array): { bytes: Uint8Array; extension: string } {
  const body = new Uint8Array(72 + dib.byteLength);
  const view = new DataView(body.buffer);
  // Offsets are relative to the record start, which has an 8-byte prefix.
  view.setUint32(40, 80, true);
  view.setUint32(44, 40, true);
  view.setUint32(48, 120, true);
  view.setUint32(52, 4, true);
  body.set(dib, 72);
  return {
    bytes: concat([
      emfRecord(1, new Uint8Array(80)),
      emfRecord(81, body),
      emfRecord(14, new Uint8Array(12)),
    ]),
    extension: "emf",
  };
}

function createBitmapWmf(dib: Uint8Array): { bytes: Uint8Array; extension: string } {
  const placeable = new Uint8Array(22);
  new DataView(placeable.buffer).setUint32(0, 0x9ac6cdd7, true);
  const header = new Uint8Array(18);
  const headerView = new DataView(header.buffer);
  headerView.setUint16(0, 1, true);
  headerView.setUint16(2, 9, true);
  const parameters = new Uint8Array(20 + dib.byteLength);
  parameters.set(dib, 20);
  return {
    bytes: concat([
      placeable,
      header,
      wmfRecord(0x0b41, parameters),
      wmfRecord(0x0000, new Uint8Array(0)),
    ]),
    extension: "wmf",
  };
}

function emfRecord(type: number, body: Uint8Array): Uint8Array {
  const record = new Uint8Array(8 + body.byteLength);
  const view = new DataView(record.buffer);
  view.setUint32(0, type, true);
  view.setUint32(4, record.byteLength, true);
  record.set(body, 8);
  return record;
}

function wmfRecord(type: number, parameters: Uint8Array): Uint8Array {
  const record = new Uint8Array(6 + parameters.byteLength);
  const view = new DataView(record.buffer);
  view.setUint32(0, record.byteLength / 2, true);
  view.setUint16(4, type, true);
  record.set(parameters, 6);
  return record;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const result = new Uint8Array(parts.reduce((total, part) => total + part.byteLength, 0));
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.byteLength;
  }
  return result;
}

function flatten(root: MockElement): MockElement[] {
  return [root, ...root.children.flatMap(flatten)];
}
