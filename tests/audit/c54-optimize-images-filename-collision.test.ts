/**
 * optimizeImages() BMP→PNG rename must not clobber a distinct existing media part.
 *
 * The new filename was derived purely by swapping the extension (image1.bmp →
 * image1.png). When the package already contains a *different* image named
 * image1.png (legal per OPC; non-Word producers don't share Word's global
 * numbering), updateEntryFilename pointed two distinct images at the same media
 * path and saveImages() wrote both to word/media/image1.png — last write wins,
 * silently destroying one image. The fix resolves the converted name to a unique
 * filename before renaming, and ImageManager.updateEntryFilename now throws if a
 * different entry already owns the target name.
 */

import * as zlib from 'zlib';
import { Document } from '../../src/core/Document';
import { Image } from '../../src/elements/Image';

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
  const crcValue = crc32(Buffer.concat([typeBuffer, data]));
  const crcBuffer = Buffer.alloc(4);
  crcBuffer.writeUInt32BE(crcValue, 0);
  return Buffer.concat([length, typeBuffer, data, crcBuffer]);
}

/** Build a well-compressed PNG that optimizeImage() will leave as-is (no savings). */
function buildOptimalPng(width: number, height: number): Buffer {
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
      rawData[y * rowSize + 1 + x * 3] = 0x12;
      rawData[y * rowSize + 1 + x * 3 + 1] = 0x34;
      rawData[y * rowSize + 1 + x * 3 + 2] = 0x56;
    }
  }
  const compressed = zlib.deflateSync(rawData, { level: 9 });
  return Buffer.concat([
    PNG_SIGNATURE,
    buildChunk('IHDR', ihdr),
    buildChunk('IDAT', compressed),
    buildChunk('IEND', Buffer.alloc(0)),
  ]);
}

/** Build a 24-bit BMP (optimizeImage converts BMP → PNG). */
function buildBmp24(width: number, height: number): Buffer {
  const rowSize = Math.ceil((width * 3) / 4) * 4;
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
      buf[offset] = 0xff;
      buf[offset + 1] = 0x00;
      buf[offset + 2] = 0x00;
    }
  }
  return buf;
}

describe('optimizeImages BMP→PNG filename collision (data-loss guard)', () => {
  it('does not rename a converted BMP onto a distinct existing PNG part', async () => {
    const doc = Document.create();
    const imageManager = doc.getImageManager();
    const relManager = doc.getRelationshipManager();

    // A distinct PNG that already occupies the base name "image1".
    const pngImage = await Image.fromBuffer(buildOptimalPng(8, 8), {
      width: 914400,
      height: 914400,
    });
    const pngRel = relManager.addImage('media/image1.png');
    imageManager.registerImage(pngImage, pngRel.getId(), 'image1.png');

    // A BMP that will convert to PNG; its naive new name is image1.png — collision.
    const bmpImage = await Image.fromBuffer(buildBmp24(20, 20), {
      width: 914400,
      height: 914400,
    });
    const bmpRel = relManager.addImage('media/image1.bmp');
    imageManager.registerImage(bmpImage, bmpRel.getId(), 'image1.bmp');

    await doc.optimizeImages();

    const entries = imageManager.getAllImages();
    const filenames = entries.map((e) => e.filename);

    // The pre-existing PNG keeps its name; the converted BMP gets a fresh,
    // distinct name — no two entries share a filename.
    expect(filenames).toContain('image1.png');
    const uniqueNames = new Set(filenames);
    expect(uniqueNames.size).toBe(filenames.length);

    // The converted image is a PNG and is NOT named image1.png (that's the other one).
    const converted = entries.find((e) => e.image === bmpImage);
    expect(converted).toBeDefined();
    expect(converted!.filename).toMatch(/\.png$/);
    expect(converted!.filename).not.toBe('image1.png');

    // The converted relationship target points at the fresh name, not the
    // colliding one.
    const convertedRel = relManager.getRelationship(bmpRel.getId());
    expect(convertedRel!.getTarget()).toContain(converted!.filename);
    expect(convertedRel!.getTarget()).not.toMatch(/\.bmp$/);

    doc.dispose();
  });

  it('ImageManager.updateEntryFilename throws when the target name is owned by a different image', async () => {
    const doc = Document.create();
    const imageManager = doc.getImageManager();
    const relManager = doc.getRelationshipManager();

    const pngImage = await Image.fromBuffer(buildOptimalPng(8, 8), {
      width: 914400,
      height: 914400,
    });
    const pngRel = relManager.addImage('media/image1.png');
    imageManager.registerImage(pngImage, pngRel.getId(), 'image1.png');

    const otherImage = await Image.fromBuffer(buildBmp24(10, 10), {
      width: 914400,
      height: 914400,
    });
    const otherRel = relManager.addImage('media/image1.bmp');
    imageManager.registerImage(otherImage, otherRel.getId(), 'image1.bmp');

    expect(() => imageManager.updateEntryFilename(otherImage, 'image1.png')).toThrow(
      /already in use/i
    );

    doc.dispose();
  });
});
