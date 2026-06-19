/**
 * Per ECMA-376 §20.1.7.6, wp:extent and a:ext define the pre-rotation bounding
 * box; a:xfrm/@rot rotates the shape about its center. rotate(90/270) must NOT
 * swap width/height — doing so stretches a non-square bitmap into the swapped
 * box before rotating, distorting the rendered image.
 */
import { Document } from '../../src/core/Document';
import { Image } from '../../src/elements/Image';
import { ZipHandler } from '../../src/zip/ZipHandler';

function createTestImageBuffer(): Buffer {
  // 1x1 transparent PNG
  return Buffer.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);
}

const WIDTH = 914400; // 1 inch
const HEIGHT = 457200; // 0.5 inch — non-square so a swap would be visible

describe('Image.rotate(90/270) keeps pre-rotation extents', () => {
  it('does not swap width/height for 90 degree rotation', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer(), {
      width: WIDTH,
      height: HEIGHT,
    });

    image.rotate(90);

    expect(image.getRotation()).toBe(90);
    expect(image.getWidth()).toBe(WIDTH);
    expect(image.getHeight()).toBe(HEIGHT);
  });

  it('does not swap width/height for 270 degree rotation', async () => {
    const image = await Image.fromBuffer(createTestImageBuffer(), {
      width: WIDTH,
      height: HEIGHT,
    });

    image.rotate(270);

    expect(image.getRotation()).toBe(270);
    expect(image.getWidth()).toBe(WIDTH);
    expect(image.getHeight()).toBe(HEIGHT);
  });

  it('emits unrotated extents with rot in saved XML and round-trips them', async () => {
    const doc = Document.create();
    try {
      const image = await Image.fromBuffer(createTestImageBuffer(), {
        width: WIDTH,
        height: HEIGHT,
      });
      image.rotate(90);
      doc.addImage(image);

      const buffer = await doc.toBuffer();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const docXml = zip.getFileAsString('word/document.xml')!;

      expect(docXml).toMatch(new RegExp(`<wp:extent cx="${WIDTH}" cy="${HEIGHT}"\\s*/?>`));
      expect(docXml).toMatch(new RegExp(`<a:ext cx="${WIDTH}" cy="${HEIGHT}"\\s*/?>`));
      expect(docXml).toContain('rot="5400000"');

      const doc2 = await Document.loadFromBuffer(buffer);
      try {
        const images = doc2.getImages();
        expect(images.length).toBe(1);
        expect(images[0]!.image.getRotation()).toBe(90);
        expect(images[0]!.image.getWidth()).toBe(WIDTH);
        expect(images[0]!.image.getHeight()).toBe(HEIGHT);
      } finally {
        doc2.dispose();
      }
    } finally {
      doc.dispose();
    }
  });
});
