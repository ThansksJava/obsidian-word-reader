export interface ExtractedBitmap {
  bytes: Uint8Array;
  mimeType: string;
}

const BI_BITFIELDS = 3;
const BI_JPEG = 4;
const BI_PNG = 5;
const BITMAP_FILE_HEADER_SIZE = 14;

const EMR_HEADER = 1;
const EMR_EOF = 14;
const EMR_STRETCHDIBITS = 81;
// EMF records that only change state or carry comments. Any other record draws
// something, and such metafiles are left to the unsupported-image placeholder.
const EMF_STATE_RECORDS = new Set([
  EMR_HEADER, EMR_EOF, 9, 10, 11, 12, 13, 17, 18, 19, 20, 21, 22, 24, 25, 30,
  33, 34, 35, 36, 37, 38, 39, 40, 58, 70, 75, 82, 95, 98, 115,
]);

const META_EOF = 0x0000;
const META_DIBSTRETCHBLT = 0x0b41;
const META_STRETCHDIB = 0x0f43;
const WMF_PLACEABLE_KEY = 0x9ac6cdd7;
const WMF_PLACEABLE_HEADER_SIZE = 22;
const WMF_STATE_RECORDS = new Set([
  META_EOF, 0x001e, 0x0035, 0x00f7, 0x0102, 0x0103, 0x0104, 0x0105, 0x0106,
  0x0107, 0x0127, 0x012d, 0x012e, 0x01f0, 0x0201, 0x0209, 0x020b, 0x020c,
  0x020d, 0x020e, 0x0234, 0x02fa, 0x02fb, 0x02fc, 0x0416, 0x0626,
]);

/**
 * Returns the single bitmap a Windows metafile wraps, converted to a format
 * browsers decode natively. Metafiles that contain vector drawing records or
 * several bitmaps return null.
 */
export function extractMetafileBitmap(
  bytes: Uint8Array,
  extension: string,
): ExtractedBitmap | null {
  try {
    switch (extension) {
      case "emf":
        return extractEmfBitmap(bytes);
      case "wmf":
        return extractWmfBitmap(bytes);
      default:
        return null;
    }
  } catch (error) {
    if (error instanceof RangeError) {
      return null;
    }
    throw error;
  }
}

export function isMetafileExtension(extension: string): boolean {
  return extension === "emf" || extension === "wmf";
}

function extractEmfBitmap(bytes: Uint8Array): ExtractedBitmap | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (view.getUint32(0, true) !== EMR_HEADER) {
    return null;
  }
  let bitmap: ExtractedBitmap | null = null;
  let offset = 0;
  while (offset + 8 <= bytes.byteLength) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || offset + size > bytes.byteLength) {
      return null;
    }
    if (type === EMR_STRETCHDIBITS) {
      if (bitmap) {
        return null;
      }
      const bitmapInfoOffset = view.getUint32(offset + 48, true);
      const bitmapInfoSize = view.getUint32(offset + 52, true);
      const bitsOffset = view.getUint32(offset + 56, true);
      const bitsSize = view.getUint32(offset + 60, true);
      if (
        bitmapInfoOffset + bitmapInfoSize > size ||
        bitsOffset + bitsSize > size
      ) {
        return null;
      }
      bitmap = deviceIndependentBitmap(
        bytes.subarray(offset + bitmapInfoOffset, offset + bitmapInfoOffset + bitmapInfoSize),
        bytes.subarray(offset + bitsOffset, offset + bitsOffset + bitsSize),
      );
      if (!bitmap) {
        return null;
      }
    } else if (!EMF_STATE_RECORDS.has(type)) {
      return null;
    }
    if (type === EMR_EOF) {
      break;
    }
    offset += size;
  }
  return bitmap;
}

function extractWmfBitmap(bytes: Uint8Array): ExtractedBitmap | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset =
    view.getUint32(0, true) === WMF_PLACEABLE_KEY ? WMF_PLACEABLE_HEADER_SIZE : 0;
  const headerWords = view.getUint16(offset + 2, true);
  if (headerWords < 9) {
    return null;
  }
  offset += headerWords * 2;
  let bitmap: ExtractedBitmap | null = null;
  while (offset + 6 <= bytes.byteLength) {
    const size = view.getUint32(offset, true) * 2;
    const type = view.getUint16(offset + 4, true);
    if (size < 6 || offset + size > bytes.byteLength) {
      return null;
    }
    if (type === META_DIBSTRETCHBLT || type === META_STRETCHDIB) {
      if (bitmap) {
        return null;
      }
      // Fixed parameters precede the packed DIB: raster operation, optional
      // color usage, and eight 16-bit source and destination coordinates.
      const dibOffset = offset + 6 + 4 + (type === META_STRETCHDIB ? 2 : 0) + 16;
      if (dibOffset >= offset + size) {
        return null;
      }
      bitmap = packedDeviceIndependentBitmap(bytes.subarray(dibOffset, offset + size));
      if (!bitmap) {
        return null;
      }
    } else if (!WMF_STATE_RECORDS.has(type)) {
      return null;
    }
    if (type === META_EOF) {
      break;
    }
    offset += size;
  }
  return bitmap;
}

/** Converts a packed DIB (header, color table and bits) to a BMP file. */
function packedDeviceIndependentBitmap(dib: Uint8Array): ExtractedBitmap | null {
  const view = new DataView(dib.buffer, dib.byteOffset, dib.byteLength);
  const headerSize = view.getUint32(0, true);
  let colorTableSize: number;
  if (headerSize === 12) {
    const bitCount = view.getUint16(10, true);
    colorTableSize = bitCount <= 8 ? (1 << bitCount) * 3 : 0;
  } else if (headerSize >= 40) {
    const bitCount = view.getUint16(14, true);
    const compression = view.getUint32(16, true);
    const colorsUsed = view.getUint32(32, true);
    const masks = compression === BI_BITFIELDS && headerSize === 40 ? 12 : 0;
    colorTableSize =
      masks + (colorsUsed || (bitCount <= 8 ? 1 << bitCount : 0)) * 4;
  } else {
    return null;
  }
  const bitmapInfoSize = headerSize + colorTableSize;
  if (bitmapInfoSize >= dib.byteLength) {
    return null;
  }
  return deviceIndependentBitmap(
    dib.subarray(0, bitmapInfoSize),
    dib.subarray(bitmapInfoSize),
  );
}

function deviceIndependentBitmap(
  bitmapInfo: Uint8Array,
  bits: Uint8Array,
): ExtractedBitmap | null {
  if (bitmapInfo.byteLength < 12 || bits.byteLength === 0) {
    return null;
  }
  const view = new DataView(bitmapInfo.buffer, bitmapInfo.byteOffset, bitmapInfo.byteLength);
  const headerSize = view.getUint32(0, true);
  const compression = headerSize >= 40 ? view.getUint32(16, true) : 0;
  if (compression === BI_JPEG) {
    return { bytes: bits.slice(), mimeType: "image/jpeg" };
  }
  if (compression === BI_PNG) {
    return { bytes: bits.slice(), mimeType: "image/png" };
  }
  const fileSize = BITMAP_FILE_HEADER_SIZE + bitmapInfo.byteLength + bits.byteLength;
  const result = new Uint8Array(fileSize);
  const header = new DataView(result.buffer);
  header.setUint8(0, 0x42);
  header.setUint8(1, 0x4d);
  header.setUint32(2, fileSize, true);
  header.setUint32(10, BITMAP_FILE_HEADER_SIZE + bitmapInfo.byteLength, true);
  result.set(bitmapInfo, BITMAP_FILE_HEADER_SIZE);
  result.set(bits, BITMAP_FILE_HEADER_SIZE + bitmapInfo.byteLength);
  return { bytes: result, mimeType: "image/bmp" };
}
