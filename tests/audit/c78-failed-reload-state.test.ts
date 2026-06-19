/**
 * Regression: a failed ZipHandler.load / loadFromBuffer must leave the handler
 * in a deterministically empty state, not serving the previously loaded archive.
 *
 * Pre-fix, writer.clear() ran only after the reader load succeeded, so a second
 * load that threw left the writer holding the prior archive and the reader's
 * `loaded` flag stale — getFile/hasFile/save then silently returned old content.
 */

import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';
import { promises as fs } from 'fs';
import * as path from 'path';

describe('C78: failed reload leaves handler empty, not stale', () => {
  const testDir = path.join(__dirname, '..', 'temp-c78');
  const validFile = path.join(testDir, 'valid.docx');
  const invalidFile = path.join(testDir, 'invalid.docx');

  beforeEach(async () => {
    await fs.mkdir(testDir, { recursive: true });
  });

  afterEach(async () => {
    try {
      await fs.rm(testDir, { recursive: true, force: true });
    } catch {
      // Ignore cleanup errors
    }
  });

  async function makeValidDocxBuffer(): Promise<Buffer> {
    const builder = new ZipHandler();
    builder.addFile(DOCX_PATHS.CONTENT_TYPES, '<?xml version="1.0"?>');
    builder.addFile(DOCX_PATHS.RELS, '<?xml version="1.0"?>');
    builder.addFile(DOCX_PATHS.DOCUMENT, '<document>Original</document>');
    return builder.toBuffer();
  }

  test('failed loadFromBuffer after a successful load exposes no files', async () => {
    const handler = new ZipHandler();
    await handler.loadFromBuffer(await makeValidDocxBuffer());

    expect(handler.hasFile(DOCX_PATHS.DOCUMENT)).toBe(true);

    // Second load with a non-ZIP buffer must throw.
    await expect(handler.loadFromBuffer(Buffer.from('not a zip archive'))).rejects.toThrow();

    // The handler must not still serve the previously loaded archive.
    expect(handler.hasFile(DOCX_PATHS.DOCUMENT)).toBe(false);
    expect(handler.getFile(DOCX_PATHS.DOCUMENT)).toBeUndefined();
    expect(handler.getFileCount()).toBe(0);
    expect(handler.isLoaded()).toBe(false);
    expect(handler.getMode()).toBe('write');
  });

  test('failed load() after a successful load exposes no files', async () => {
    await fs.writeFile(validFile, await makeValidDocxBuffer());
    await fs.writeFile(invalidFile, 'This is not a ZIP file');

    const handler = new ZipHandler();
    await handler.load(validFile);
    expect(handler.hasFile(DOCX_PATHS.DOCUMENT)).toBe(true);

    await expect(handler.load(invalidFile)).rejects.toThrow();

    expect(handler.hasFile(DOCX_PATHS.DOCUMENT)).toBe(false);
    expect(handler.getFileCount()).toBe(0);
    expect(handler.isLoaded()).toBe(false);
    expect(handler.getMode()).toBe('write');
  });
});
