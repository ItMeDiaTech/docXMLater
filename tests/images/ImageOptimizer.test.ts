/**
 * Unit tests for ImageOptimizer
 *
 * Tests PNG re-compression and BMP → PNG conversion using synthetic test images.
 */

import * as zlib from 'zlib';
import { optimizePng, convertBmpToPng, optimizeImage } from '../../src/images/ImageOptimizer';

// =============================================================================
// Test Helpers — Synthetic Image Builders
// =============================================================================

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

/** CRC-32 for building valid PNG chunks in tests */
function crc32(buf: Buffer): number {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c;
  }
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc = table[(crc ^ buf[i]!) & 0xff]! ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function buildChunk(type: string, data: Buffer): Buffer {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);
  const typeBuffer = Buffer.from(type, 'ascii');
  const crcInput = Buffer.concat([typeBuffer, data]);
  const crcValue = crc32(crcInput);
  const crcBuffer = Buffer.alloc(4);
  crcBuffer.writeUInt32BE(crcValue, 0);
  return Buffer.concat([length, typeBuffer, data, crcBuffer]);
}

/** Build a minimal 1x1 RGBA PNG (red pixel) */
function buildMinimalPng(): Buffer {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(1, 0); // width
  ihdr.writeUInt32BE(1, 4); // height
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace

  // Raw pixel data: filter byte (0) + R(FF) G(00) B(00) A(FF)
  const rawData = Buffer.from([0, 0xff, 0x00, 0x00, 0xff]);
  const compressed = zlib.deflateSync(rawData, { level: 6 });

  return Buffer.concat([
    PNG_SIGNATURE,
    buildChunk('IHDR', ihdr),
    buildChunk('IDAT', compressed),
    buildChunk('IEND', Buffer.alloc(0)),
  ]);
}

/** Build a larger PNG with suboptimal compression (level 1) */
function buildSuboptimalPng(width: number, height: number): Buffer {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 2; // RGB
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  // Raw pixel data: for each row, filter byte (0) + RGB pixels (solid red)
  const rowSize = 1 + width * 3;
  const rawData = Buffer.alloc(height * rowSize);
  for (let y = 0; y < height; y++) {
    rawData[y * rowSize] = 0; // filter: None
    for (let x = 0; x < width; x++) {
      rawData[y * rowSize + 1 + x * 3] = 0xff; // R
      rawData[y * rowSize + 1 + x * 3 + 1] = 0x00; // G
      rawData[y * rowSize + 1 + x * 3 + 2] = 0x00; // B
    }
  }

  // Compress at level 1 (suboptimal)
  const compressed = zlib.deflateSync(rawData, { level: 1 });

  return Buffer.concat([
    PNG_SIGNATURE,
    buildChunk('IHDR', ihdr),
    buildChunk('IDAT', compressed),
    buildChunk('IEND', Buffer.alloc(0)),
  ]);
}

/** Build a PNG with metadata chunks (tEXt, iCCP) */
function buildPngWithMetadata(width: number, height: number): Buffer {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2; // RGB
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const rowSize = 1 + width * 3;
  const rawData = Buffer.alloc(height * rowSize);
  for (let y = 0; y < height; y++) {
    rawData[y * rowSize] = 0;
    for (let x = 0; x < width; x++) {
      rawData[y * rowSize + 1 + x * 3] = 0x80;
      rawData[y * rowSize + 1 + x * 3 + 1] = 0x80;
      rawData[y * rowSize + 1 + x * 3 + 2] = 0x80;
    }
  }

  const compressed = zlib.deflateSync(rawData, { level: 1 });

  // Build tEXt chunk: "Software\0TestSuite"
  const textData = Buffer.from('Software\0TestSuite', 'latin1');
  // Build a fake iCCP chunk
  const iccpData = Buffer.from('sRGB\0\0' + 'x'.repeat(100), 'latin1');

  return Buffer.concat([
    PNG_SIGNATURE,
    buildChunk('IHDR', ihdr),
    buildChunk('tEXt', textData),
    buildChunk('iCCP', iccpData),
    buildChunk('IDAT', compressed),
    buildChunk('IEND', Buffer.alloc(0)),
  ]);
}

/** Build a minimal 24-bit BMP (2x2 pixels) */
function buildBmp24(width: number, height: number): Buffer {
  const bytesPerPixel = 3;
  const rowSize = Math.ceil((width * bytesPerPixel) / 4) * 4;
  const pixelDataSize = rowSize * height;
  const fileSize = 54 + pixelDataSize;

  const buf = Buffer.alloc(fileSize);

  // File header (14 bytes)
  buf[0] = 0x42;
  buf[1] = 0x4d; // BM
  buf.writeUInt32LE(fileSize, 2);
  buf.writeUInt32LE(0, 6); // reserved
  buf.writeUInt32LE(54, 10); // pixel data offset

  // DIB header (40 bytes - BITMAPINFOHEADER)
  buf.writeUInt32LE(40, 14); // header size
  buf.writeInt32LE(width, 18);
  buf.writeInt32LE(height, 22); // positive = bottom-up
  buf.writeUInt16LE(1, 26); // color planes
  buf.writeUInt16LE(24, 28); // bits per pixel
  buf.writeUInt32LE(0, 30); // compression (none)
  buf.writeUInt32LE(pixelDataSize, 34);
  buf.writeInt32LE(2835, 38); // X pixels/meter (~72 DPI)
  buf.writeInt32LE(2835, 42); // Y pixels/meter
  buf.writeUInt32LE(0, 46); // colors used
  buf.writeUInt32LE(0, 50); // important colors

  // Pixel data (BGR, bottom-up): fill with solid blue (B=FF, G=00, R=00)
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const offset = 54 + y * rowSize + x * 3;
      buf[offset] = 0xff; // B
      buf[offset + 1] = 0x00; // G
      buf[offset + 2] = 0x00; // R
    }
  }

  return buf;
}

/** Build a 32-bit RGBA BMP */
function buildBmp32(width: number, height: number): Buffer {
  const bytesPerPixel = 4;
  const rowSize = width * bytesPerPixel; // 32-bit rows are always 4-byte aligned
  const pixelDataSize = rowSize * height;
  // For BI_BITFIELDS, we need 12 extra bytes for RGB masks after the header
  const headerSize = 54 + 12; // file header(14) + DIB header(40) + masks(12)
  const fileSize = headerSize + pixelDataSize;

  const buf = Buffer.alloc(fileSize);

  // File header (14 bytes)
  buf[0] = 0x42;
  buf[1] = 0x4d;
  buf.writeUInt32LE(fileSize, 2);
  buf.writeUInt32LE(0, 6);
  buf.writeUInt32LE(headerSize, 10); // pixel data offset after masks

  // DIB header (40 bytes)
  buf.writeUInt32LE(40, 14);
  buf.writeInt32LE(width, 18);
  buf.writeInt32LE(height, 22);
  buf.writeUInt16LE(1, 26);
  buf.writeUInt16LE(32, 28); // 32 bpp
  buf.writeUInt32LE(3, 30); // BI_BITFIELDS
  buf.writeUInt32LE(pixelDataSize, 34);
  buf.writeInt32LE(2835, 38);
  buf.writeInt32LE(2835, 42);
  buf.writeUInt32LE(0, 46);
  buf.writeUInt32LE(0, 50);

  // Color masks (12 bytes) for BI_BITFIELDS
  buf.writeUInt32LE(0x00ff0000, 54); // R mask
  buf.writeUInt32LE(0x0000ff00, 58); // G mask
  buf.writeUInt32LE(0x000000ff, 62); // B mask

  // Pixel data (BGRA, bottom-up): fill with semi-transparent green
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const offset = headerSize + y * rowSize + x * 4;
      buf[offset] = 0x00; // B
      buf[offset + 1] = 0xff; // G
      buf[offset + 2] = 0x00; // R
      buf[offset + 3] = 0x80; // A (128 = semi-transparent)
    }
  }

  return buf;
}

/** Build a BMP with RLE compression (unsupported variant) */
function buildRleCompressedBmp(): Buffer {
  const buf = Buffer.alloc(66);

  // File header
  buf[0] = 0x42;
  buf[1] = 0x4d;
  buf.writeUInt32LE(66, 2);
  buf.writeUInt32LE(0, 6);
  buf.writeUInt32LE(54, 10);

  // DIB header
  buf.writeUInt32LE(40, 14);
  buf.writeInt32LE(2, 18);
  buf.writeInt32LE(2, 22);
  buf.writeUInt16LE(1, 26);
  buf.writeUInt16LE(8, 28); // 8-bit indexed
  buf.writeUInt32LE(1, 30); // RLE8 compression
  buf.writeUInt32LE(12, 34);

  return buf;
}

// =============================================================================
// PNG Re-compression Tests
// =============================================================================

describe('optimizePng', () => {
  it('should re-compress a suboptimal PNG to be smaller', () => {
    const original = buildSuboptimalPng(100, 100);
    const optimized = optimizePng(original);

    expect(optimized).not.toBeNull();
    expect(optimized!.length).toBeLessThan(original.length);

    // Verify it's still a valid PNG
    expect(optimized!.subarray(0, 8).equals(PNG_SIGNATURE)).toBe(true);

    // Verify dimensions are preserved (IHDR chunk starts at offset 8+4+4=16)
    const width = optimized!.readUInt32BE(16);
    const height = optimized!.readUInt32BE(20);
    expect(width).toBe(100);
    expect(height).toBe(100);
  });

  it('should preserve pixel data through re-compression', () => {
    const original = buildSuboptimalPng(10, 10);
    const optimized = optimizePng(original);

    expect(optimized).not.toBeNull();

    // Decompress both IDATs and compare pixel data
    const originalIdat = extractIdatData(original);
    const optimizedIdat = extractIdatData(optimized!);

    const originalPixels = zlib.inflateSync(originalIdat);
    const optimizedPixels = zlib.inflateSync(optimizedIdat);

    expect(originalPixels.equals(optimizedPixels)).toBe(true);
  });

  it('should return null for an already-optimal PNG when called via optimizeImage', () => {
    // Build a PNG already compressed at level 9 with no metadata
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(2, 0);
    ihdr.writeUInt32BE(2, 4);
    ihdr[8] = 8;
    ihdr[9] = 2;

    const rawData = Buffer.alloc(2 * (1 + 2 * 3));
    const compressed = zlib.deflateSync(rawData, { level: 9 });

    const optimal = Buffer.concat([
      PNG_SIGNATURE,
      buildChunk('IHDR', ihdr),
      buildChunk('IDAT', compressed),
      buildChunk('IEND', Buffer.alloc(0)),
    ]);

    // optimizeImage returns null when optimized size >= original
    const result = optimizeImage(optimal, 'png');
    expect(result).toBeNull();
  });

  it('should work with tiny PNGs', () => {
    const tiny = buildMinimalPng();
    const optimized = optimizePng(tiny);

    // For tiny PNGs, the optimized version may be the same size or even larger
    // Just verify it doesn't crash and returns a valid PNG
    if (optimized) {
      expect(optimized.subarray(0, 8).equals(PNG_SIGNATURE)).toBe(true);
    }
  });

  it('should strip tEXt and iCCP metadata chunks', () => {
    const withMeta = buildPngWithMetadata(50, 50);
    const optimized = optimizePng(withMeta);

    expect(optimized).not.toBeNull();

    // Verify no tEXt or iCCP chunks remain
    const chunks = parsePngChunkTypes(optimized!);
    expect(chunks).not.toContain('tEXt');
    expect(chunks).not.toContain('iCCP');

    // Should only contain essential chunks
    for (const chunk of chunks) {
      expect(['IHDR', 'PLTE', 'tRNS', 'IDAT', 'IEND']).toContain(chunk);
    }
  });

  it('should preserve PLTE and tRNS chunks for indexed PNGs', () => {
    // Build a palette-based PNG (color type 3)
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(4, 0);
    ihdr.writeUInt32BE(4, 4);
    ihdr[8] = 8;
    ihdr[9] = 3; // indexed color

    // 2-entry palette: red and blue
    const plte = Buffer.from([0xff, 0x00, 0x00, 0x00, 0x00, 0xff]);
    // tRNS for palette transparency
    const trns = Buffer.from([0xff, 0x80]);

    // Raw data: 4 rows, 4 pixels each, filter 0
    const rawData = Buffer.alloc(4 * (1 + 4));
    for (let y = 0; y < 4; y++) {
      rawData[y * 5] = 0;
      for (let x = 0; x < 4; x++) {
        rawData[y * 5 + 1 + x] = (x + y) % 2;
      }
    }
    const compressed = zlib.deflateSync(rawData, { level: 1 });

    const png = Buffer.concat([
      PNG_SIGNATURE,
      buildChunk('IHDR', ihdr),
      buildChunk('PLTE', plte),
      buildChunk('tRNS', trns),
      buildChunk('IDAT', compressed),
      buildChunk('IEND', Buffer.alloc(0)),
    ]);

    const optimized = optimizePng(png);
    expect(optimized).not.toBeNull();

    const chunks = parsePngChunkTypes(optimized!);
    expect(chunks).toContain('PLTE');
    expect(chunks).toContain('tRNS');
  });

  it('should return null for non-PNG data', () => {
    const jpeg = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10]);
    expect(optimizePng(jpeg)).toBeNull();
  });

  it('should return null for empty buffer', () => {
    expect(optimizePng(Buffer.alloc(0))).toBeNull();
  });

  it('should handle PNG with multiple IDAT chunks', () => {
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(20, 0);
    ihdr.writeUInt32BE(20, 4);
    ihdr[8] = 8;
    ihdr[9] = 2;

    const rawData = Buffer.alloc(20 * (1 + 20 * 3));
    const compressed = zlib.deflateSync(rawData, { level: 1 });

    // Split into 3 IDAT chunks
    const chunk1 = compressed.subarray(0, Math.floor(compressed.length / 3));
    const chunk2 = compressed.subarray(
      Math.floor(compressed.length / 3),
      Math.floor((2 * compressed.length) / 3)
    );
    const chunk3 = compressed.subarray(Math.floor((2 * compressed.length) / 3));

    const png = Buffer.concat([
      PNG_SIGNATURE,
      buildChunk('IHDR', ihdr),
      buildChunk('IDAT', chunk1),
      buildChunk('IDAT', chunk2),
      buildChunk('IDAT', chunk3),
      buildChunk('IEND', Buffer.alloc(0)),
    ]);

    const optimized = optimizePng(png);
    expect(optimized).not.toBeNull();

    // Verify pixel data is identical
    const originalPixels = zlib.inflateSync(compressed);
    const optimizedIdat = extractIdatData(optimized!);
    const optimizedPixels = zlib.inflateSync(optimizedIdat);
    expect(originalPixels.equals(optimizedPixels)).toBe(true);
  });
});

// =============================================================================
// BMP → PNG Conversion Tests
// =============================================================================

describe('convertBmpToPng', () => {
  it('should convert a 24-bit BMP to valid PNG', () => {
    const bmp = buildBmp24(4, 4);
    const png = convertBmpToPng(bmp);

    expect(png).not.toBeNull();
    expect(png!.subarray(0, 8).equals(PNG_SIGNATURE)).toBe(true);

    // Verify dimensions
    const width = png!.readUInt32BE(16);
    const height = png!.readUInt32BE(20);
    expect(width).toBe(4);
    expect(height).toBe(4);

    // Verify color type is RGB (2)
    expect(png![25]).toBe(2);
  });

  it('should convert a 32-bit RGBA BMP to valid PNG with alpha', () => {
    const bmp = buildBmp32(4, 4);
    const png = convertBmpToPng(bmp);

    expect(png).not.toBeNull();

    // Verify dimensions
    const width = png!.readUInt32BE(16);
    const height = png!.readUInt32BE(20);
    expect(width).toBe(4);
    expect(height).toBe(4);

    // Verify color type is RGBA (6)
    expect(png![25]).toBe(6);

    // Verify alpha channel is preserved by decompressing IDAT
    const idatData = extractIdatData(png!);
    const pixels = zlib.inflateSync(idatData);

    // First pixel of first row (after filter byte): R=0, G=FF, B=0, A=80
    expect(pixels[1]).toBe(0x00); // R
    expect(pixels[2]).toBe(0xff); // G
    expect(pixels[3]).toBe(0x00); // B
    expect(pixels[4]).toBe(0x80); // A (semi-transparent)
  });

  it('should correctly flip BMP bottom-up rows to PNG top-down', () => {
    // Build a 2x2 BMP with distinct colors per row
    const bmp = buildBmp24(2, 2);

    // Manually set pixels: bottom row (y=0 in BMP) = red, top row (y=1) = blue
    // BMP offset: 54 + row * rowSize + x * 3
    const rowSize = Math.ceil((2 * 3) / 4) * 4; // 8 bytes (padded)

    // Bottom row (y=0 in BMP, last row visually): red (BGR = 00,00,FF)
    bmp[54 + 0] = 0x00;
    bmp[54 + 1] = 0x00;
    bmp[54 + 2] = 0xff; // pixel (0,0)
    bmp[54 + 3] = 0x00;
    bmp[54 + 4] = 0x00;
    bmp[54 + 5] = 0xff; // pixel (1,0)

    // Top row (y=1 in BMP, first row visually): blue (BGR = FF,00,00)
    bmp[54 + rowSize + 0] = 0xff;
    bmp[54 + rowSize + 1] = 0x00;
    bmp[54 + rowSize + 2] = 0x00;
    bmp[54 + rowSize + 3] = 0xff;
    bmp[54 + rowSize + 4] = 0x00;
    bmp[54 + rowSize + 5] = 0x00;

    const png = convertBmpToPng(bmp);
    expect(png).not.toBeNull();

    const idatData = extractIdatData(png!);
    const pixels = zlib.inflateSync(idatData);

    // PNG is top-down: first row should be blue (from BMP's top row y=1)
    // Row 0: filter(0) + RGB pixels
    const row0Start = 1; // skip filter byte
    expect(pixels[row0Start]).toBe(0x00); // R (from BMP B=FF → but wait, it's BGR!)
    // Actually: BMP top row (y=1) is BGR = FF,00,00 → RGB = 00,00,FF (blue)
    expect(pixels[row0Start]).toBe(0x00); // R
    expect(pixels[row0Start + 1]).toBe(0x00); // G
    expect(pixels[row0Start + 2]).toBe(0xff); // B → blue

    // Row 1: filter(0) + RGB pixels
    const row1Start = 1 + 2 * 3 + 1; // second row filter byte
    // BMP bottom row (y=0) is BGR = 00,00,FF → RGB = FF,00,00 (red)
    expect(pixels[row1Start]).toBe(0xff); // R
    expect(pixels[row1Start + 1]).toBe(0x00); // G
    expect(pixels[row1Start + 2]).toBe(0x00); // B → red
  });

  it('should return null for RLE-compressed BMP', () => {
    const rle = buildRleCompressedBmp();
    expect(convertBmpToPng(rle)).toBeNull();
  });

  it('should return null for 8-bit indexed BMP', () => {
    const buf = Buffer.alloc(58);
    buf[0] = 0x42;
    buf[1] = 0x4d;
    buf.writeUInt32LE(58, 2);
    buf.writeUInt32LE(54, 10);
    buf.writeUInt32LE(40, 14);
    buf.writeInt32LE(2, 18);
    buf.writeInt32LE(2, 22);
    buf.writeUInt16LE(1, 26);
    buf.writeUInt16LE(8, 28); // 8-bit
    buf.writeUInt32LE(0, 30);
    expect(convertBmpToPng(buf)).toBeNull();
  });

  it('should return null for non-BMP data', () => {
    expect(convertBmpToPng(Buffer.from([0x89, 0x50, 0x4e, 0x47]))).toBeNull();
  });

  it('should return null for empty buffer', () => {
    expect(convertBmpToPng(Buffer.alloc(0))).toBeNull();
  });

  it('should produce a PNG significantly smaller than the BMP', () => {
    // A 100x100 24-bit BMP should be much larger than the resulting PNG
    const bmp = buildBmp24(100, 100);
    const png = convertBmpToPng(bmp);

    expect(png).not.toBeNull();
    // BMP for 100x100 solid color ~= 30054 bytes, PNG should be much smaller
    expect(png!.length).toBeLessThan(bmp.length / 5);
  });
});

// =============================================================================
// Router (optimizeImage) Tests
// =============================================================================

describe('optimizeImage', () => {
  it('should optimize PNG files', () => {
    const png = buildSuboptimalPng(50, 50);
    const result = optimizeImage(png, 'png');

    expect(result).not.toBeNull();
    expect(result!.newExtension).toBe('png');
    expect(result!.data.length).toBeLessThan(png.length);
  });

  it('should convert BMP files to PNG', () => {
    const bmp = buildBmp24(10, 10);
    const result = optimizeImage(bmp, 'bmp');

    expect(result).not.toBeNull();
    expect(result!.newExtension).toBe('png');
    // Verify the result is valid PNG
    expect(result!.data.subarray(0, 8).equals(PNG_SIGNATURE)).toBe(true);
  });

  it('should return null for JPEG', () => {
    const jpeg = Buffer.from([0xff, 0xd8, 0xff, 0xe0]);
    expect(optimizeImage(jpeg, 'jpeg')).toBeNull();
    expect(optimizeImage(jpeg, 'jpg')).toBeNull();
  });

  it('should return null for EMF', () => {
    expect(optimizeImage(Buffer.alloc(100), 'emf')).toBeNull();
  });

  it('should return null for WMF', () => {
    expect(optimizeImage(Buffer.alloc(100), 'wmf')).toBeNull();
  });

  it('should return null for SVG', () => {
    expect(optimizeImage(Buffer.from('<svg></svg>'), 'svg')).toBeNull();
  });

  it('should return null for GIF', () => {
    expect(optimizeImage(Buffer.alloc(100), 'gif')).toBeNull();
  });

  it('should handle extension case-insensitively', () => {
    const bmp = buildBmp24(4, 4);
    const result = optimizeImage(bmp, 'BMP');
    expect(result).not.toBeNull();
    expect(result!.newExtension).toBe('png');
  });
});

// =============================================================================
// Helpers
// =============================================================================

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

/** Parse all chunk types from a PNG buffer */
function parsePngChunkTypes(png: Buffer): string[] {
  const types: string[] = [];
  let offset = 8;
  while (offset + 12 <= png.length) {
    const length = png.readUInt32BE(offset);
    const type = png.subarray(offset + 4, offset + 8).toString('ascii');
    types.push(type);
    offset += 12 + length;
    if (type === 'IEND') break;
  }
  return types;
}
