/**
 * Tests for ZIP/XML resource limits (availability hardening for untrusted DOCX input).
 *
 * These build small synthetic archives in-memory with JSZip — never a weaponized
 * "zip bomb" — and assert that breaching a (deliberately tiny) configured budget
 * raises a typed {@link ResourceLimitError}, while a normal small DOCX still loads
 * with the generous defaults.
 */

import JSZip from 'jszip';
import { promises as fs } from 'fs';
import * as os from 'os';
import * as path from 'path';
import { ZipReader } from '../../src/zip/ZipReader';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { ResourceLimitError } from '../../src/zip/errors';
import { DEFAULT_SIZE_LIMITS } from '../../src/zip/types';

/**
 * Builds a minimal, structurally-valid DOCX buffer (the three required parts).
 */
async function buildMinimalDocx(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ' +
      'Target="word/document.xml"/></Relationships>'
  );
  zip.file(
    'word/document.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>'
  );
  return zip.generateAsync({ type: 'nodebuffer' });
}

describe('ZIP resource limits', () => {
  describe('SizeLimitOptions defaults', () => {
    test('expose generous single-sourced defaults', () => {
      expect(DEFAULT_SIZE_LIMITS.maxTotalUncompressedMB).toBe(300);
      expect(DEFAULT_SIZE_LIMITS.maxEntryUncompressedMB).toBe(150);
      expect(DEFAULT_SIZE_LIMITS.maxEntryCount).toBe(2000);
      expect(DEFAULT_SIZE_LIMITS.maxCompressionRatio).toBe(200);
      // Pre-existing fields preserved
      expect(DEFAULT_SIZE_LIMITS.maxSizeMB).toBe(150);
      expect(DEFAULT_SIZE_LIMITS.warningSizeMB).toBe(50);
    });
  });

  describe('entry-count budget', () => {
    test('rejects an archive whose entry count exceeds maxEntryCount', async () => {
      const zip = new JSZip();
      for (let i = 0; i < 12; i++) {
        zip.file(`part-${i}.bin`, `entry ${i}`);
      }
      const buffer = await zip.generateAsync({ type: 'nodebuffer' });

      const reader = new ZipReader();
      try {
        await expect(
          reader.loadFromBuffer(buffer, { validate: false, sizeLimits: { maxEntryCount: 5 } })
        ).rejects.toThrow(ResourceLimitError);
        await expect(
          reader.loadFromBuffer(buffer, { validate: false, sizeLimits: { maxEntryCount: 5 } })
        ).rejects.toThrow(/maxEntryCount/);
      } finally {
        reader.clear();
      }
    });

    test('boundary: entry count exactly at maxEntryCount loads; limit+1 throws [T9]', async () => {
      // Build an archive with a known number of non-directory entries.
      const entryCount = 4;
      const zip = new JSZip();
      for (let i = 0; i < entryCount; i++) {
        zip.file(`part-${i}.bin`, `entry ${i}`);
      }
      const buffer = await zip.generateAsync({ type: 'nodebuffer' });

      // At the limit (count === maxEntryCount): inclusive, must load.
      const atLimit = new ZipReader();
      try {
        await atLimit.loadFromBuffer(buffer, {
          validate: false,
          sizeLimits: { maxEntryCount: entryCount },
        });
        expect(atLimit.isLoaded()).toBe(true);
        expect(atLimit.getFilePaths()).toHaveLength(entryCount);
      } finally {
        atLimit.clear();
      }

      // One over the limit (count === maxEntryCount + 1, i.e. maxEntryCount = count - 1): rejected.
      const overLimit = new ZipReader();
      try {
        await expect(
          overLimit.loadFromBuffer(buffer, {
            validate: false,
            sizeLimits: { maxEntryCount: entryCount - 1 },
          })
        ).rejects.toThrow(ResourceLimitError);
      } finally {
        overLimit.clear();
      }
    });
  });

  describe('compressed-size budget (maxSizeMB)', () => {
    test('ZipHandler.loadFromBuffer throws ResourceLimitError when maxSizeMB is exceeded [T3]', async () => {
      // A normal small archive is a few hundred bytes; a deliberately tiny maxSizeMB
      // (~104 bytes) guarantees the compressed-size guard fires for it.
      const buffer = await buildMinimalDocx();
      expect(buffer.length).toBeGreaterThan(0.0001 * 1024 * 1024);

      const handler = new ZipHandler();
      try {
        await expect(
          handler.loadFromBuffer(buffer, { validate: false, sizeLimits: { maxSizeMB: 0.0001 } })
        ).rejects.toThrow(ResourceLimitError);
        // The message still names the breached limit so it stays actionable.
        await expect(
          handler.loadFromBuffer(buffer, { validate: false, sizeLimits: { maxSizeMB: 0.0001 } })
        ).rejects.toThrow(/maximum supported size/);
      } finally {
        handler.clear();
      }
    });
  });

  describe('total-uncompressed budget (primary guard)', () => {
    test('rejects when the running decompressed total exceeds maxTotalUncompressedMB', async () => {
      const zip = new JSZip();
      // ~100 KB of text, split across entries so the in-loop accumulator is exercised.
      zip.file('a.txt', 'a'.repeat(60 * 1024));
      zip.file('b.txt', 'b'.repeat(60 * 1024));
      const buffer = await zip.generateAsync({ type: 'nodebuffer' });

      const reader = new ZipReader();
      try {
        await expect(
          reader.loadFromBuffer(buffer, {
            validate: false,
            // ~52 KB budget — well below the ~120 KB total.
            sizeLimits: { maxTotalUncompressedMB: 0.05 },
          })
        ).rejects.toThrow(ResourceLimitError);
      } finally {
        reader.clear();
      }
    });
  });

  describe('per-entry uncompressed budget', () => {
    test('rejects when a single entry exceeds maxEntryUncompressedMB', async () => {
      const zip = new JSZip();
      zip.file('big.txt', 'x'.repeat(100 * 1024));
      const buffer = await zip.generateAsync({ type: 'nodebuffer' });

      const reader = new ZipReader();
      try {
        await expect(
          reader.loadFromBuffer(buffer, {
            validate: false,
            sizeLimits: { maxEntryUncompressedMB: 0.05 },
          })
        ).rejects.toThrow(/maxEntryUncompressedMB/);
      } finally {
        reader.clear();
      }
    });
  });

  describe('compression-ratio budget', () => {
    test('rejects a sizable entry whose deflate ratio exceeds maxCompressionRatio', async () => {
      const zip = new JSZip();
      // 2 MB of identical bytes deflates to a few KB → ratio in the hundreds.
      zip.file('compressible.bin', Buffer.alloc(2 * 1024 * 1024, 0x41));
      const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });

      const reader = new ZipReader();
      try {
        await expect(
          reader.loadFromBuffer(buffer, {
            validate: false,
            // Generous size budgets so only the ratio guard can fire.
            sizeLimits: { maxCompressionRatio: 10 },
          })
        ).rejects.toThrow(/maxCompressionRatio/);
      } finally {
        reader.clear();
      }
    });
  });

  describe('normal load with defaults', () => {
    test('a small valid DOCX loads unchanged under default limits', async () => {
      const buffer = await buildMinimalDocx();
      const reader = new ZipReader();
      try {
        await reader.loadFromBuffer(buffer);
        expect(reader.isLoaded()).toBe(true);
        expect(reader.hasFile('word/document.xml')).toBe(true);
        expect(reader.getFileAsString('word/document.xml')).toContain('Hello');
      } finally {
        reader.clear();
      }
    });

    test('a multi-entry small DOCX stays under the default entry-count budget', async () => {
      const zip = new JSZip();
      for (let i = 0; i < 50; i++) {
        zip.file(`word/media/image${i}.bin`, Buffer.from([0x00, i & 0xff]));
      }
      // Required parts so structure validation also passes.
      const base = await buildMinimalDocx();
      const baseZip = await JSZip.loadAsync(base);
      for (const name of Object.keys(baseZip.files)) {
        if (!baseZip.files[name]!.dir) {
          zip.file(name, await baseZip.files[name]!.async('nodebuffer'));
        }
      }
      const buffer = await zip.generateAsync({ type: 'nodebuffer' });

      const reader = new ZipReader();
      try {
        await reader.loadFromBuffer(buffer);
        expect(reader.isLoaded()).toBe(true);
      } finally {
        reader.clear();
      }
    });
  });

  describe('error surfacing through ZipHandler facade', () => {
    test('loadFromBuffer surfaces ResourceLimitError, not CorruptedArchiveError', async () => {
      const zip = new JSZip();
      for (let i = 0; i < 6; i++) {
        zip.file(`f${i}.bin`, 'data');
      }
      const buffer = await zip.generateAsync({ type: 'nodebuffer' });

      const handler = new ZipHandler();
      await expect(
        handler.loadFromBuffer(buffer, { validate: false, sizeLimits: { maxEntryCount: 2 } })
      ).rejects.toThrow(ResourceLimitError);
      handler.clear();
    });

    test('loadFromFile surfaces ResourceLimitError instead of rewrapping as FileOperationError', async () => {
      const zip = new JSZip();
      for (let i = 0; i < 6; i++) {
        zip.file(`f${i}.bin`, 'data');
      }
      const buffer = await zip.generateAsync({ type: 'nodebuffer' });

      const tmpFile = path.join(os.tmpdir(), `docxmlater-resource-limit-${Date.now()}.docx`);
      await fs.writeFile(tmpFile, buffer);
      try {
        const reader = new ZipReader();
        await expect(
          reader.loadFromFile(tmpFile, { validate: false, sizeLimits: { maxEntryCount: 2 } })
        ).rejects.toThrow(ResourceLimitError);
        reader.clear();
      } finally {
        await fs.rm(tmpFile, { force: true });
      }
    });
  });
});
