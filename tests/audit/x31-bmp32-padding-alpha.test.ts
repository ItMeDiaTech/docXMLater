/**
 * convertBmpToPng — 32-bpp BMP alpha/channel handling
 *
 * BI_RGB (compression=0) 32-bpp BMPs define the 4th byte per pixel as reserved
 * padding (commonly zeroed). It must not be copied into the PNG alpha channel
 * blindly, or the converted image becomes fully transparent. BI_BITFIELDS
 * (compression=3) BMPs carry explicit channel masks at offset 54 that must be
 * honored; non-byte-aligned masks cannot be converted losslessly.
 */

import * as zlib from 'zlib';
import { convertBmpToPng } from '../../src/images/ImageOptimizer';

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

/** Extract concatenated IDAT data from a PNG buffer */
function extractIdatData(png: Buffer): Buffer {
  const idatBuffers: Buffer[] = [];
  let offset = 8;
  while (offset + 12 <= png.length) {
    const length = png.readUInt32BE(offset);
    const type = png.subarray(offset + 4, offset + 8).toString('ascii');
    if (type === 'IDAT') {
      idatBuffers.push(png.subarray(offset + 8, offset + 8 + length));
    }
    offset += 12 + length;
    if (type === 'IEND') break;
  }
  return Buffer.concat(idatBuffers);
}

/** Build a 32-bpp BI_RGB BMP (no masks, pixel data at offset 54), all pixels identical BGRX bytes */
function buildBmp32BiRgb(
  width: number,
  height: number,
  pixelBytes: [number, number, number, number]
): Buffer {
  const rowSize = width * 4;
  const pixelDataSize = rowSize * height;
  const fileSize = 54 + pixelDataSize;
  const buf = Buffer.alloc(fileSize);

  // File header (14 bytes)
  buf[0] = 0x42;
  buf[1] = 0x4d;
  buf.writeUInt32LE(fileSize, 2);
  buf.writeUInt32LE(54, 10); // pixel data offset

  // DIB header (40 bytes - BITMAPINFOHEADER)
  buf.writeUInt32LE(40, 14);
  buf.writeInt32LE(width, 18);
  buf.writeInt32LE(height, 22);
  buf.writeUInt16LE(1, 26);
  buf.writeUInt16LE(32, 28); // 32 bpp
  buf.writeUInt32LE(0, 30); // BI_RGB (uncompressed)
  buf.writeUInt32LE(pixelDataSize, 34);

  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const off = 54 + y * rowSize + x * 4;
      buf[off] = pixelBytes[0];
      buf[off + 1] = pixelBytes[1];
      buf[off + 2] = pixelBytes[2];
      buf[off + 3] = pixelBytes[3];
    }
  }

  return buf;
}

/** Build a 32-bpp BI_BITFIELDS BMP with explicit R/G/B masks, all pixels identical raw bytes */
function buildBmp32Bitfields(
  width: number,
  height: number,
  masks: [number, number, number],
  pixelBytes: [number, number, number, number]
): Buffer {
  const rowSize = width * 4;
  const pixelDataSize = rowSize * height;
  const headerSize = 54 + 12; // file header(14) + DIB header(40) + masks(12)
  const fileSize = headerSize + pixelDataSize;
  const buf = Buffer.alloc(fileSize);

  buf[0] = 0x42;
  buf[1] = 0x4d;
  buf.writeUInt32LE(fileSize, 2);
  buf.writeUInt32LE(headerSize, 10);

  buf.writeUInt32LE(40, 14);
  buf.writeInt32LE(width, 18);
  buf.writeInt32LE(height, 22);
  buf.writeUInt16LE(1, 26);
  buf.writeUInt16LE(32, 28);
  buf.writeUInt32LE(3, 30); // BI_BITFIELDS
  buf.writeUInt32LE(pixelDataSize, 34);

  buf.writeUInt32LE(masks[0], 54); // R mask
  buf.writeUInt32LE(masks[1], 58); // G mask
  buf.writeUInt32LE(masks[2], 62); // B mask

  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const off = headerSize + y * rowSize + x * 4;
      buf[off] = pixelBytes[0];
      buf[off + 1] = pixelBytes[1];
      buf[off + 2] = pixelBytes[2];
      buf[off + 3] = pixelBytes[3];
    }
  }

  return buf;
}

/** Decode the first pixel (RGBA) of a converted PNG */
function firstPixel(png: Buffer): [number, number, number, number] {
  const pixels = zlib.inflateSync(extractIdatData(png));
  // Row layout: filter byte (0) + RGBA pixels
  return [pixels[1]!, pixels[2]!, pixels[3]!, pixels[4]!];
}

describe('convertBmpToPng 32-bpp alpha and channel masks', () => {
  it('should emit opaque pixels for BI_RGB 32-bpp with zeroed padding bytes', () => {
    // Solid red, 4th byte zeroed (typical reserved padding)
    const bmp = buildBmp32BiRgb(4, 4, [0x00, 0x00, 0xff, 0x00]);
    const png = convertBmpToPng(bmp);

    expect(png).not.toBeNull();
    expect(png!.subarray(0, 8).equals(PNG_SIGNATURE)).toBe(true);

    const [r, g, b, a] = firstPixel(png!);
    expect(r).toBe(0xff);
    expect(g).toBe(0x00);
    expect(b).toBe(0x00);
    expect(a).toBe(0xff); // opaque — padding must not become alpha=0
  });

  it('should preserve alpha for BI_RGB 32-bpp when the 4th byte carries real values', () => {
    const bmp = buildBmp32BiRgb(4, 4, [0x00, 0xff, 0x00, 0x80]);
    const png = convertBmpToPng(bmp);

    expect(png).not.toBeNull();
    const [r, g, b, a] = firstPixel(png!);
    expect(r).toBe(0x00);
    expect(g).toBe(0xff);
    expect(b).toBe(0x00);
    expect(a).toBe(0x80);
  });

  it('should route channels by BI_BITFIELDS masks instead of assuming BGRA', () => {
    // RGBA byte order: R in byte 0, G in byte 1, B in byte 2
    const masks: [number, number, number] = [0x000000ff, 0x0000ff00, 0x00ff0000];
    // Pure red, fully opaque, stored as R,G,B,A bytes
    const bmp = buildBmp32Bitfields(4, 4, masks, [0xff, 0x00, 0x00, 0xff]);
    const png = convertBmpToPng(bmp);

    expect(png).not.toBeNull();
    const [r, g, b, a] = firstPixel(png!);
    expect(r).toBe(0xff);
    expect(g).toBe(0x00);
    expect(b).toBe(0x00);
    expect(a).toBe(0xff);
  });

  it('should still convert BI_BITFIELDS with standard BGRA masks as before', () => {
    const masks: [number, number, number] = [0x00ff0000, 0x0000ff00, 0x000000ff];
    const bmp = buildBmp32Bitfields(4, 4, masks, [0x00, 0xff, 0x00, 0x80]);
    const png = convertBmpToPng(bmp);

    expect(png).not.toBeNull();
    const [r, g, b, a] = firstPixel(png!);
    expect(r).toBe(0x00);
    expect(g).toBe(0xff);
    expect(b).toBe(0x00);
    expect(a).toBe(0x80);
  });

  it('should return null for BI_BITFIELDS masks that are not byte-aligned', () => {
    // 16-bit-wide red mask cannot map losslessly onto an 8-bit PNG channel
    const masks: [number, number, number] = [0xffff0000, 0x0000ff00, 0x000000ff];
    const bmp = buildBmp32Bitfields(4, 4, masks, [0x00, 0x00, 0xff, 0xff]);
    expect(convertBmpToPng(bmp)).toBeNull();
  });

  it('should return null for BI_BITFIELDS masks that overlap the same byte', () => {
    const masks: [number, number, number] = [0x00ff0000, 0x00ff0000, 0x000000ff];
    const bmp = buildBmp32Bitfields(4, 4, masks, [0x00, 0x00, 0xff, 0xff]);
    expect(convertBmpToPng(bmp)).toBeNull();
  });
});
