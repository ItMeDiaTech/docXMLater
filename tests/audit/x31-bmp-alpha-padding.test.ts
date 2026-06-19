/**
 * Regression tests for convertBmpToPng 32-bpp alpha/mask handling.
 *
 * - BI_RGB (compression=0) 32-bpp: the 4th byte per pixel is reserved padding,
 *   typically zeroed. It must not be copied into PNG alpha verbatim, or the
 *   converted image becomes fully transparent.
 * - BI_BITFIELDS (compression=3) 32-bpp: the three channel masks at offset 54
 *   must be honored; non-standard byte orders must be routed correctly and
 *   non-byte-aligned masks rejected (null) to keep the conversion lossless.
 */

import * as zlib from 'zlib';
import { convertBmpToPng } from '../../src/images/ImageOptimizer';

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

interface Bmp32Options {
  width: number;
  height: number;
  compression: 0 | 3;
  /** [R, G, B] masks; required when compression === 3 */
  masks?: [number, number, number];
  /** Per-pixel 4 raw bytes (little-endian pixel layout), row 0 = top row */
  pixelBytes: (x: number, y: number) => [number, number, number, number];
}

/** Build a 32-bpp BMP with explicit compression mode and raw pixel bytes (bottom-up storage) */
function buildBmp32(opts: Bmp32Options): Buffer {
  const { width, height, compression, masks, pixelBytes } = opts;
  const rowSize = width * 4; // 32-bit rows are always 4-byte aligned
  const maskBytes = compression === 3 ? 12 : 0;
  const pixelDataOffset = 54 + maskBytes;
  const fileSize = pixelDataOffset + rowSize * height;

  const buf = Buffer.alloc(fileSize);

  // File header (14 bytes)
  buf[0] = 0x42;
  buf[1] = 0x4d;
  buf.writeUInt32LE(fileSize, 2);
  buf.writeUInt32LE(pixelDataOffset, 10);

  // BITMAPINFOHEADER (40 bytes)
  buf.writeUInt32LE(40, 14);
  buf.writeInt32LE(width, 18);
  buf.writeInt32LE(height, 22);
  buf.writeUInt16LE(1, 26);
  buf.writeUInt16LE(32, 28);
  buf.writeUInt32LE(compression, 30);
  buf.writeUInt32LE(rowSize * height, 34);

  if (compression === 3 && masks) {
    buf.writeUInt32LE(masks[0], 54); // R mask
    buf.writeUInt32LE(masks[1], 58); // G mask
    buf.writeUInt32LE(masks[2], 62); // B mask
  }

  // Pixel data, stored bottom-up
  for (let y = 0; y < height; y++) {
    const bmpRow = height - 1 - y; // storage row for logical (top-down) row y
    for (let x = 0; x < width; x++) {
      const offset = pixelDataOffset + bmpRow * rowSize + x * 4;
      const [b0, b1, b2, b3] = pixelBytes(x, y);
      buf[offset] = b0;
      buf[offset + 1] = b1;
      buf[offset + 2] = b2;
      buf[offset + 3] = b3;
    }
  }

  return buf;
}

/** Decode the raw RGBA scanlines of a filter-0 PNG produced by convertBmpToPng */
function decodePngRgba(png: Buffer): { width: number; height: number; pixels: Buffer } {
  expect(png.subarray(0, 8).equals(PNG_SIGNATURE)).toBe(true);

  const width = png.readUInt32BE(16);
  const height = png.readUInt32BE(20);
  expect(png[25]).toBe(6); // color type RGBA

  // Walk chunks to collect IDAT data
  const idatParts: Buffer[] = [];
  let offset = 8;
  while (offset + 12 <= png.length) {
    const length = png.readUInt32BE(offset);
    const type = png.subarray(offset + 4, offset + 8).toString('ascii');
    if (type === 'IDAT') {
      idatParts.push(png.subarray(offset + 8, offset + 8 + length));
    }
    offset += 12 + length;
    if (type === 'IEND') break;
  }

  const raw = zlib.inflateSync(Buffer.concat(idatParts));
  const stride = 1 + width * 4;
  const pixels = Buffer.alloc(width * height * 4);
  for (let y = 0; y < height; y++) {
    expect(raw[y * stride]).toBe(0); // filter: None
    raw.copy(pixels, y * width * 4, y * stride + 1, (y + 1) * stride);
  }
  return { width, height, pixels };
}

function pixelAt(decoded: { width: number; pixels: Buffer }, x: number, y: number): number[] {
  const o = (y * decoded.width + x) * 4;
  return [...decoded.pixels.subarray(o, o + 4)];
}

describe('convertBmpToPng 32-bpp alpha and channel-mask handling', () => {
  it('treats zeroed BI_RGB padding bytes as opaque, not transparent', () => {
    // BGRX layout with X (padding) = 0x00 everywhere
    const bmp = buildBmp32({
      width: 2,
      height: 2,
      compression: 0,
      pixelBytes: (x, y) => {
        if (x === 0 && y === 0) return [0x00, 0x00, 0xff, 0x00]; // red
        if (x === 1 && y === 0) return [0x00, 0xff, 0x00, 0x00]; // green
        if (x === 0 && y === 1) return [0xff, 0x00, 0x00, 0x00]; // blue
        return [0x10, 0x20, 0x30, 0x00]; // arbitrary color
      },
    });

    const png = convertBmpToPng(bmp);
    expect(png).not.toBeNull();

    const decoded = decodePngRgba(png!);
    expect(decoded.width).toBe(2);
    expect(decoded.height).toBe(2);
    expect(pixelAt(decoded, 0, 0)).toEqual([0xff, 0x00, 0x00, 0xff]);
    expect(pixelAt(decoded, 1, 0)).toEqual([0x00, 0xff, 0x00, 0xff]);
    expect(pixelAt(decoded, 0, 1)).toEqual([0x00, 0x00, 0xff, 0xff]);
    expect(pixelAt(decoded, 1, 1)).toEqual([0x30, 0x20, 0x10, 0xff]);
  });

  it('preserves BI_RGB 4th bytes as alpha when any pixel has a non-zero value', () => {
    const bmp = buildBmp32({
      width: 2,
      height: 1,
      compression: 0,
      pixelBytes: (x) => (x === 0 ? [0x00, 0x00, 0xff, 0x80] : [0x00, 0xff, 0x00, 0x00]),
    });

    const png = convertBmpToPng(bmp);
    expect(png).not.toBeNull();

    const decoded = decodePngRgba(png!);
    expect(pixelAt(decoded, 0, 0)).toEqual([0xff, 0x00, 0x00, 0x80]);
    expect(pixelAt(decoded, 1, 0)).toEqual([0x00, 0xff, 0x00, 0x00]);
  });

  it('routes channels by BI_BITFIELDS masks for non-standard (RGBA-order) layouts', () => {
    // Masks place R in byte 0, G in byte 1, B in byte 2 — opposite of BGRA
    const bmp = buildBmp32({
      width: 2,
      height: 1,
      compression: 3,
      masks: [0x000000ff, 0x0000ff00, 0x00ff0000],
      pixelBytes: (x) =>
        x === 0
          ? [0xff, 0x00, 0x00, 0xff] // R=255 → red
          : [0x00, 0x00, 0xff, 0x80], // B=255 → blue, semi-transparent
    });

    const png = convertBmpToPng(bmp);
    expect(png).not.toBeNull();

    const decoded = decodePngRgba(png!);
    expect(pixelAt(decoded, 0, 0)).toEqual([0xff, 0x00, 0x00, 0xff]);
    expect(pixelAt(decoded, 1, 0)).toEqual([0x00, 0x00, 0xff, 0x80]);
  });

  it('still converts BI_BITFIELDS with the standard BGRA masks', () => {
    const bmp = buildBmp32({
      width: 1,
      height: 1,
      compression: 3,
      masks: [0x00ff0000, 0x0000ff00, 0x000000ff],
      pixelBytes: () => [0x00, 0xff, 0x00, 0x80], // BGRA: green, semi-transparent
    });

    const png = convertBmpToPng(bmp);
    expect(png).not.toBeNull();

    const decoded = decodePngRgba(png!);
    expect(pixelAt(decoded, 0, 0)).toEqual([0x00, 0xff, 0x00, 0x80]);
  });

  it('returns null for BI_BITFIELDS masks that are not byte-aligned 8-bit channels', () => {
    // 5:5:5-style masks cannot map losslessly onto PNG 8-bit channels
    const bmp = buildBmp32({
      width: 1,
      height: 1,
      compression: 3,
      masks: [0x7c000000, 0x03e00000, 0x001f0000],
      pixelBytes: () => [0x00, 0x00, 0x00, 0x00],
    });

    expect(convertBmpToPng(bmp)).toBeNull();
  });

  it('returns null for BI_BITFIELDS files too short to contain the masks', () => {
    const bmp = buildBmp32({
      width: 1,
      height: 1,
      compression: 3,
      masks: [0x00ff0000, 0x0000ff00, 0x000000ff],
      pixelBytes: () => [0x00, 0x00, 0x00, 0x00],
    });
    // Truncate just before the masks; keep the declared sizes intact
    expect(convertBmpToPng(bmp.subarray(0, 54))).toBeNull();
  });
});
