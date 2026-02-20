/**
 * ImageOptimizer - Lossless image optimization for DOCX documents
 *
 * Two techniques using only Node.js built-in zlib (zero dependencies):
 * 1. PNG re-compression — Re-compress IDAT chunks at zlib level 9 + strip metadata
 * 2. BMP → PNG conversion — Lossless format change, typically 10-50x smaller
 */

import * as zlib from 'zlib';

// =============================================================================
// Types
// =============================================================================

export interface ImageOptimizationResult {
  optimizedCount: number;
  totalSavedBytes: number;
}

// =============================================================================
// CRC-32 (IEEE polynomial)
// =============================================================================

const CRC_TABLE: Uint32Array = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[n] = c;
  }
  return table;
})();

function crc32(buf: Buffer): number {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) {
    crc = CRC_TABLE[(crc ^ buf[i]!) & 0xFF]! ^ (crc >>> 8);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

// =============================================================================
// PNG Chunk Builder
// =============================================================================

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

/**
 * Builds a PNG chunk: [4-byte length][4-byte type][data][4-byte CRC]
 * CRC covers type + data bytes
 */
function buildPngChunk(type: string, data: Buffer): Buffer {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length, 0);

  const typeBuffer = Buffer.from(type, 'ascii');

  const crcInput = Buffer.concat([typeBuffer, data]);
  const crcValue = crc32(crcInput);
  const crcBuffer = Buffer.alloc(4);
  crcBuffer.writeUInt32BE(crcValue, 0);

  return Buffer.concat([length, typeBuffer, data, crcBuffer]);
}

// =============================================================================
// PNG Re-compression
// =============================================================================

/** Chunks that must be preserved for correct PNG rendering */
const ESSENTIAL_CHUNK_TYPES = new Set(['IHDR', 'PLTE', 'tRNS', 'IDAT', 'IEND']);

/**
 * Re-compresses a PNG buffer at zlib level 9 and strips non-essential metadata chunks.
 * Returns the optimized buffer, or null if the PNG is invalid/cannot be optimized.
 * The caller should compare sizes to decide whether to use the result.
 */
export function optimizePng(buffer: Buffer): Buffer | null {
  // Verify PNG signature
  if (buffer.length < 8 || !buffer.subarray(0, 8).equals(PNG_SIGNATURE)) {
    return null;
  }

  // Parse chunks
  let ihdrData: Buffer | null = null;
  let plteData: Buffer | null = null;
  let trnsData: Buffer | null = null;
  const idatBuffers: Buffer[] = [];

  let offset = 8; // Skip signature
  while (offset + 12 <= buffer.length) {
    const chunkLength = buffer.readUInt32BE(offset);
    const chunkType = buffer.subarray(offset + 4, offset + 8).toString('ascii');

    // Bounds check: length + type(4) + data + crc(4) = 12 + chunkLength
    if (offset + 12 + chunkLength > buffer.length) break;

    const chunkData = buffer.subarray(offset + 8, offset + 8 + chunkLength);

    if (chunkType === 'IHDR') {
      ihdrData = Buffer.from(chunkData);
    } else if (chunkType === 'PLTE') {
      plteData = Buffer.from(chunkData);
    } else if (chunkType === 'tRNS') {
      trnsData = Buffer.from(chunkData);
    } else if (chunkType === 'IDAT') {
      idatBuffers.push(chunkData);
    }
    // Non-essential chunks (tEXt, iTXt, zTXt, iCCP, sRGB, gAMA, cHRM, tIME, pHYs, etc.)
    // are intentionally discarded

    offset += 12 + chunkLength;

    // Stop after IEND
    if (chunkType === 'IEND') break;
  }

  if (!ihdrData || idatBuffers.length === 0) return null;

  // Decompress all IDAT data
  const combinedIdat = Buffer.concat(idatBuffers);
  let decompressed: Buffer;
  try {
    decompressed = zlib.inflateSync(combinedIdat);
  } catch {
    return null; // Corrupted IDAT data
  }

  // Re-compress at maximum level
  const recompressed = zlib.deflateSync(decompressed, { level: 9 });

  // Rebuild PNG: signature + IHDR + [PLTE] + [tRNS] + IDAT + IEND
  const resultParts: Buffer[] = [Buffer.from(PNG_SIGNATURE)];

  resultParts.push(buildPngChunk('IHDR', ihdrData));
  if (plteData) resultParts.push(buildPngChunk('PLTE', plteData));
  if (trnsData) resultParts.push(buildPngChunk('tRNS', trnsData));
  resultParts.push(buildPngChunk('IDAT', recompressed));
  resultParts.push(buildPngChunk('IEND', Buffer.alloc(0)));

  return Buffer.concat(resultParts);
}

// =============================================================================
// BMP → PNG Conversion
// =============================================================================

/**
 * Converts a BMP buffer to PNG format. Lossless conversion.
 * Supports 24-bit (RGB) and 32-bit (RGBA) uncompressed BMPs.
 * Returns null for unsupported variants (indexed, 16-bit, RLE-compressed).
 */
export function convertBmpToPng(buffer: Buffer): Buffer | null {
  // Minimum BMP size: 14 (file header) + 40 (BITMAPINFOHEADER) = 54 bytes
  if (buffer.length < 54 || buffer[0] !== 0x42 || buffer[1] !== 0x4D) {
    return null;
  }

  // Parse BMP file header (14 bytes)
  const pixelDataOffset = buffer.readUInt32LE(10);

  // Parse DIB header (BITMAPINFOHEADER)
  const dibHeaderSize = buffer.readUInt32LE(14);
  if (dibHeaderSize < 40) return null; // Only support BITMAPINFOHEADER and larger

  const width = buffer.readInt32LE(18);
  const height = buffer.readInt32LE(22);
  const bitsPerPixel = buffer.readUInt16LE(28);
  const compression = buffer.readUInt32LE(30);

  // Only support uncompressed (0) and BI_BITFIELDS (3) for 32-bit
  if (compression !== 0 && !(compression === 3 && bitsPerPixel === 32)) return null;

  // Only support 24-bit and 32-bit
  if (bitsPerPixel !== 24 && bitsPerPixel !== 32) return null;

  if (width <= 0) return null;
  const absHeight = Math.abs(height);
  if (absHeight === 0) return null;
  const isTopDown = height < 0;

  const bytesPerPixel = bitsPerPixel / 8;

  // BMP rows are padded to 4-byte boundaries
  const rowSize = Math.ceil((width * bytesPerPixel) / 4) * 4;

  // Validate buffer has enough pixel data
  if (pixelDataOffset + absHeight * rowSize > buffer.length) return null;

  // PNG color type: 2 = RGB (24-bit), 6 = RGBA (32-bit)
  const colorType = bitsPerPixel === 32 ? 6 : 2;
  const pngBytesPerPixel = bitsPerPixel === 32 ? 4 : 3;

  // Build IHDR data (13 bytes)
  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(width, 0);
  ihdrData.writeUInt32BE(absHeight, 4);
  ihdrData[8] = 8;         // bit depth
  ihdrData[9] = colorType;  // color type
  ihdrData[10] = 0;         // compression method (deflate)
  ihdrData[11] = 0;         // filter method (adaptive)
  ihdrData[12] = 0;         // interlace method (none)

  // Build raw image data: filter byte (0 = None) + pixel data per row
  const rawDataSize = absHeight * (1 + width * pngBytesPerPixel);
  const rawData = Buffer.alloc(rawDataSize);

  for (let y = 0; y < absHeight; y++) {
    // BMP stores rows bottom-up by default; top-down if height is negative
    const bmpRow = isTopDown ? y : (absHeight - 1 - y);
    const bmpRowOffset = pixelDataOffset + bmpRow * rowSize;

    const pngRowOffset = y * (1 + width * pngBytesPerPixel);
    rawData[pngRowOffset] = 0; // Filter type: None

    for (let x = 0; x < width; x++) {
      const bmpPixelOffset = bmpRowOffset + x * bytesPerPixel;
      const pngPixelOffset = pngRowOffset + 1 + x * pngBytesPerPixel;

      // Convert BGR(A) → RGB(A)
      rawData[pngPixelOffset] = buffer[bmpPixelOffset + 2]!;     // R
      rawData[pngPixelOffset + 1] = buffer[bmpPixelOffset + 1]!; // G
      rawData[pngPixelOffset + 2] = buffer[bmpPixelOffset]!;     // B

      if (bitsPerPixel === 32) {
        rawData[pngPixelOffset + 3] = buffer[bmpPixelOffset + 3]!; // A
      }
    }
  }

  // Compress pixel data
  const compressedData = zlib.deflateSync(rawData, { level: 9 });

  // Assemble PNG file
  return Buffer.concat([
    Buffer.from(PNG_SIGNATURE),
    buildPngChunk('IHDR', ihdrData),
    buildPngChunk('IDAT', compressedData),
    buildPngChunk('IEND', Buffer.alloc(0)),
  ]);
}

// =============================================================================
// Router
// =============================================================================

/**
 * Optimizes an image buffer based on its format.
 * - PNG: re-compresses at zlib level 9 and strips metadata. Only returns result if smaller.
 * - BMP: converts to PNG (always smaller than uncompressed BMP).
 * - Other formats: returns null (unsupported / cannot be losslessly optimized further).
 */
export function optimizeImage(
  buffer: Buffer,
  extension: string
): { data: Buffer; newExtension: string } | null {
  const ext = extension.toLowerCase();

  if (ext === 'png') {
    const optimized = optimizePng(buffer);
    if (!optimized || optimized.length >= buffer.length) return null;
    return { data: optimized, newExtension: 'png' };
  }

  if (ext === 'bmp') {
    const converted = convertBmpToPng(buffer);
    if (!converted) return null;
    return { data: converted, newExtension: 'png' };
  }

  return null;
}
