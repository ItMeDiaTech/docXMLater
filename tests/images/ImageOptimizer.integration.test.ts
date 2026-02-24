/**
 * Integration tests for doc.optimizeImages()
 *
 * Tests the full document-level optimization flow including:
 * - PNG re-compression in a document context
 * - BMP → PNG conversion with filename/relationship updates
 * - Mixed image types (JPEG untouched, PNG/BMP optimized)
 * - Deduplication of shared image files
 */

import * as zlib from 'zlib';
import { Document } from '../../src/core/Document';
import { Image } from '../../src/elements/Image';

// =============================================================================
// Test Helpers
// =============================================================================

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

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

/** Build a suboptimal PNG (level 1 compression + metadata) */
function buildSuboptimalPng(width: number, height: number): Buffer {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2; // RGB
  const rowSize = 1 + width * 3;
  const rawData = Buffer.alloc(height * rowSize);
  for (let y = 0; y < height; y++) {
    rawData[y * rowSize] = 0;
    for (let x = 0; x < width; x++) {
      rawData[y * rowSize + 1 + x * 3] = 0xff;
      rawData[y * rowSize + 1 + x * 3 + 1] = 0x00;
      rawData[y * rowSize + 1 + x * 3 + 2] = 0x00;
    }
  }
  const compressed = zlib.deflateSync(rawData, { level: 1 });
  const textData = Buffer.from('Software\0TestSuite', 'latin1');
  return Buffer.concat([
    PNG_SIGNATURE,
    buildChunk('IHDR', ihdr),
    buildChunk('tEXt', textData),
    buildChunk('IDAT', compressed),
    buildChunk('IEND', Buffer.alloc(0)),
  ]);
}

/** Build a 24-bit BMP */
function buildBmp24(width: number, height: number): Buffer {
  const bytesPerPixel = 3;
  const rowSize = Math.ceil((width * bytesPerPixel) / 4) * 4;
  const pixelDataSize = rowSize * height;
  const fileSize = 54 + pixelDataSize;
  const buf = Buffer.alloc(fileSize);

  buf[0] = 0x42;
  buf[1] = 0x4d;
  buf.writeUInt32LE(fileSize, 2);
  buf.writeUInt32LE(54, 10);
  buf.writeUInt32LE(40, 14);
  buf.writeInt32LE(width, 18);
  buf.writeInt32LE(height, 22);
  buf.writeUInt16LE(1, 26);
  buf.writeUInt16LE(24, 28);
  buf.writeUInt32LE(0, 30);
  buf.writeUInt32LE(pixelDataSize, 34);

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

/** Build a minimal valid JPEG buffer */
function buildMinimalJpeg(): Buffer {
  // Smallest valid JPEG: SOI + APP0 (JFIF) + minimal DQT + SOF0 + DHT + SOS + EOI
  // For testing purposes, a minimal JPEG header is sufficient
  return Buffer.from([
    0xff,
    0xd8, // SOI
    0xff,
    0xe0, // APP0
    0x00,
    0x10, // Length: 16
    0x4a,
    0x46,
    0x49,
    0x46,
    0x00, // JFIF\0
    0x01,
    0x01, // Version 1.1
    0x00, // Density units
    0x00,
    0x01, // X density
    0x00,
    0x01, // Y density
    0x00,
    0x00, // Thumbnail
    0xff,
    0xd9, // EOI
  ]);
}

// =============================================================================
// Integration Tests
// =============================================================================

describe('Document.optimizeImages() Integration', () => {
  it('should optimize a PNG image in a document', async () => {
    const doc = Document.create();
    const pngBuffer = buildSuboptimalPng(80, 80);
    const image = await Image.fromBuffer(pngBuffer, { width: 914400, height: 914400 });
    doc.addImage(image);

    const result = await doc.optimizeImages();

    expect(result.optimizedCount).toBe(1);
    expect(result.totalSavedBytes).toBeGreaterThan(0);

    // Verify the image is still a valid PNG after optimization
    const images = doc.getImages();
    expect(images.length).toBe(1);
    const optimizedData = images[0]!.image.getImageData();
    expect(optimizedData.subarray(0, 8).equals(PNG_SIGNATURE)).toBe(true);
  });

  it('should convert a BMP image to PNG', async () => {
    const doc = Document.create();
    const bmpBuffer = buildBmp24(20, 20);
    const image = await Image.fromBuffer(bmpBuffer, { width: 914400, height: 914400 });
    doc.addImage(image);

    // Verify it starts as BMP
    expect(image.getExtension()).toBe('bmp');

    const result = await doc.optimizeImages();

    expect(result.optimizedCount).toBe(1);
    expect(result.totalSavedBytes).toBeGreaterThan(0);

    // Verify image is now PNG
    const images = doc.getImages();
    expect(images.length).toBe(1);
    expect(images[0]!.image.getExtension()).toBe('png');

    // Verify filename changed
    expect(images[0]!.filename).toMatch(/\.png$/);

    // Verify the image data is valid PNG
    const data = images[0]!.image.getImageData();
    expect(data.subarray(0, 8).equals(PNG_SIGNATURE)).toBe(true);
  });

  it('should update relationship target for BMP → PNG conversion', async () => {
    const doc = Document.create();
    const bmpBuffer = buildBmp24(10, 10);
    const image = await Image.fromBuffer(bmpBuffer, { width: 914400, height: 914400 });
    doc.addImage(image);

    await doc.optimizeImages();

    // Verify the relationship target was updated
    const relManager = doc.getRelationshipManager();
    const images = doc.getImages();
    const relId = images[0]!.image.getRelationshipId();
    expect(relId).toBeDefined();

    const rel = relManager.getRelationship(relId!);
    expect(rel).toBeDefined();
    expect(rel!.getTarget()).toMatch(/\.png$/);
    expect(rel!.getTarget()).not.toMatch(/\.bmp$/);
  });

  it('should leave JPEG images untouched', async () => {
    const doc = Document.create();
    const jpegBuffer = buildMinimalJpeg();
    const image = await Image.fromBuffer(jpegBuffer, { width: 914400, height: 914400 });
    doc.addImage(image);

    const originalSize = jpegBuffer.length;
    const result = await doc.optimizeImages();

    expect(result.optimizedCount).toBe(0);
    expect(result.totalSavedBytes).toBe(0);

    // Verify JPEG is unchanged
    const images = doc.getImages();
    expect(images[0]!.image.getExtension()).toBe('jpeg');
    expect(images[0]!.image.getImageData().length).toBe(originalSize);
  });

  it('should handle mixed image types (PNG, BMP, JPEG)', async () => {
    const doc = Document.create();

    // Add a suboptimal PNG
    const pngBuffer = buildSuboptimalPng(50, 50);
    const pngImage = await Image.fromBuffer(pngBuffer, { width: 914400, height: 914400 });
    doc.addImage(pngImage);

    // Add a BMP
    const bmpBuffer = buildBmp24(20, 20);
    const bmpImage = await Image.fromBuffer(bmpBuffer, { width: 914400, height: 914400 });
    doc.addImage(bmpImage);

    // Add a JPEG
    const jpegBuffer = buildMinimalJpeg();
    const jpegImage = await Image.fromBuffer(jpegBuffer, { width: 914400, height: 914400 });
    doc.addImage(jpegImage);

    const result = await doc.optimizeImages();

    // PNG and BMP should be optimized, JPEG should not
    expect(result.optimizedCount).toBe(2);
    expect(result.totalSavedBytes).toBeGreaterThan(0);

    // Verify types after optimization
    const images = doc.getImages();
    // All three images should still be present
    expect(images.length).toBe(3);

    // Find each by checking type
    const extensions = images.map((i) => i.image.getExtension());
    expect(extensions.filter((e) => e === 'png').length).toBe(2); // Original PNG + converted BMP
    expect(extensions.filter((e) => e === 'jpeg').length).toBe(1); // JPEG unchanged
  });

  it('should save and reload a document with optimized images', async () => {
    const doc = Document.create();
    const pngBuffer = buildSuboptimalPng(40, 40);
    const image = await Image.fromBuffer(pngBuffer, { width: 914400, height: 914400 });
    doc.addImage(image);

    await doc.optimizeImages();

    // Save to buffer and reload
    const savedBuffer = await doc.toBuffer();
    const reloaded = await Document.loadFromBuffer(savedBuffer);

    // Verify image is present and valid
    const images = reloaded.getImages();
    expect(images.length).toBe(1);
    expect(images[0]!.image.getExtension()).toBe('png');

    // Ensure data loaded
    await images[0]!.image.ensureDataLoaded();
    const data = images[0]!.image.getImageData();
    expect(data.subarray(0, 8).equals(PNG_SIGNATURE)).toBe(true);
  });

  it('should save and reload a document with BMP→PNG converted images', async () => {
    const doc = Document.create();
    const bmpBuffer = buildBmp24(30, 30);
    const image = await Image.fromBuffer(bmpBuffer, { width: 914400, height: 914400 });
    doc.addImage(image);

    await doc.optimizeImages();

    // Save to buffer and reload
    const savedBuffer = await doc.toBuffer();
    const reloaded = await Document.loadFromBuffer(savedBuffer);

    const images = reloaded.getImages();
    expect(images.length).toBe(1);

    // Verify it's PNG now
    expect(images[0]!.image.getExtension()).toBe('png');

    // Verify filename
    expect(images[0]!.filename).toMatch(/\.png$/);

    // Verify dimensions preserved
    await images[0]!.image.ensureDataLoaded();
    const data = images[0]!.image.getImageData();
    expect(data.subarray(0, 8).equals(PNG_SIGNATURE)).toBe(true);

    // Read IHDR to verify dimensions
    const width = data.readUInt32BE(16);
    const height = data.readUInt32BE(20);
    expect(width).toBe(30);
    expect(height).toBe(30);
  });

  it('should return zero when no images can be optimized', async () => {
    const doc = Document.create();

    // Add only a JPEG
    const jpegBuffer = buildMinimalJpeg();
    const image = await Image.fromBuffer(jpegBuffer, { width: 914400, height: 914400 });
    doc.addImage(image);

    const result = await doc.optimizeImages();
    expect(result.optimizedCount).toBe(0);
    expect(result.totalSavedBytes).toBe(0);
  });

  it('should return zero for a document with no images', async () => {
    const doc = Document.create();
    const result = await doc.optimizeImages();
    expect(result.optimizedCount).toBe(0);
    expect(result.totalSavedBytes).toBe(0);
  });
});
