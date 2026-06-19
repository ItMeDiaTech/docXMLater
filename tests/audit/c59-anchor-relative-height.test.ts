/**
 * Per ECMA-376 CT_Anchor (dml-wordprocessingDrawing.xsd), relativeHeight is a
 * required attribute of wp:anchor. An image floated via setPosition() alone
 * (the documented anchor-less path) must still emit relativeHeight, defaulting
 * to 251658240 to match the parser fallback and the float* helpers.
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

describe('wp:anchor emits required relativeHeight without setAnchor()', () => {
  it('defaults relativeHeight to 251658240 on the setPosition()-only path', async () => {
    const doc = Document.create();
    try {
      const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
      image.setPosition({ anchor: 'page', offset: 914400 }, { anchor: 'page', offset: 914400 });
      expect(image.isFloating()).toBe(true);

      doc.addImage(image);

      const zip = new ZipHandler();
      await zip.loadFromBuffer(await doc.toBuffer());
      const docXml = zip.getFileAsString('word/document.xml')!;

      const anchorMatch = docXml.match(/<wp:anchor\b[^>]*>/);
      expect(anchorMatch).not.toBeNull();
      expect(anchorMatch![0]).toContain('relativeHeight="251658240"');
    } finally {
      doc.dispose();
    }
  });

  it('keeps an explicit relativeHeight from setAnchor()', async () => {
    const doc = Document.create();
    try {
      const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
      image.setPosition({ anchor: 'page', offset: 914400 }, { anchor: 'page', offset: 914400 });
      image.setAnchor({
        behindDoc: false,
        locked: false,
        layoutInCell: true,
        allowOverlap: false,
        relativeHeight: 42,
      });

      doc.addImage(image);

      const zip = new ZipHandler();
      await zip.loadFromBuffer(await doc.toBuffer());
      const docXml = zip.getFileAsString('word/document.xml')!;

      const anchorMatch = docXml.match(/<wp:anchor\b[^>]*>/);
      expect(anchorMatch).not.toBeNull();
      expect(anchorMatch![0]).toContain('relativeHeight="42"');
    } finally {
      doc.dispose();
    }
  });
});
