/**
 * Document.dispose() must release tracked image buffers before clearing the
 * image map.
 *
 * dispose() previously called imageManager.clear() before
 * imageManager.releaseAllImageData(); clear() empties the images map, so the
 * subsequent releaseAllImageData() iterated zero entries and Image.releaseData()
 * was never invoked. A user-held Image with path-sourced data kept its buffer
 * pinned. Swapping the order releases the buffer first.
 */

import { promises as fs } from 'fs';
import * as os from 'os';
import * as path from 'path';

import { Document } from '../../src/core/Document';
import { Image } from '../../src/elements/Image';

// Minimal valid 1x1 PNG.
const PNG_BYTES = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVQI12NgAAIABQAB' +
    'Nl7BcQAAAABJRU5ErkJggg==',
  'base64'
);

describe('Document.dispose() releases path-sourced image data', () => {
  let pngPath: string;

  beforeAll(async () => {
    pngPath = path.join(os.tmpdir(), `docxmlater-c79-${process.pid}-${Date.now()}.png`);
    await fs.writeFile(pngPath, PNG_BYTES);
  });

  afterAll(async () => {
    await fs.rm(pngPath, { force: true });
  });

  it('Image.releaseData() runs on dispose so the loaded buffer is freed', async () => {
    const doc = Document.create();

    // Path-sourced image: releaseData() only clears data for string sources.
    const image = await Image.fromFile(pngPath);
    doc.addImage(image);

    // Force the buffer into memory so there is something to release.
    await image.ensureDataLoaded();
    expect(image.getImageData().length).toBeGreaterThan(0);

    doc.dispose();

    // Pre-fix: clear() ran first, releaseAllImageData() iterated zero entries,
    // so the buffer stayed loaded and getImageData() would still return it.
    expect(() => image.getImageData()).toThrow();
  });
});
