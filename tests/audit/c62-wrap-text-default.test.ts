/**
 * Per ECMA-376 (dml-wordprocessingDrawing.xsd), wrapText is use="required" on
 * CT_WrapSquare, CT_WrapTight, and CT_WrapThrough. setWrap('square'|'tight'|
 * 'through') without a side must default wrapText to "bothSides" so the
 * emitted wp:anchor stays schema-valid. Wrap types without wrapText
 * (topAndBottom, none) must not gain the attribute.
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

async function saveFloatingImageXml(configure: (image: Image) => void): Promise<string> {
  const doc = Document.create();
  try {
    const image = await Image.fromBuffer(createTestImageBuffer(), 'png', 914400, 914400);
    image.setAnchor({
      behindDoc: false,
      locked: false,
      layoutInCell: true,
      allowOverlap: false,
      relativeHeight: 251658240,
    });
    image.setPosition({ anchor: 'page', offset: 914400 }, { anchor: 'page', offset: 914400 });
    configure(image);
    doc.addImage(image);

    const zip = new ZipHandler();
    await zip.loadFromBuffer(await doc.toBuffer());
    return zip.getFileAsString('word/document.xml')!;
  } finally {
    doc.dispose();
  }
}

describe('setWrap without a side emits required wrapText', () => {
  it('defaults wrapText to bothSides for square wrap', async () => {
    const docXml = await saveFloatingImageXml((image) => image.setWrap('square'));

    const wrapMatch = docXml.match(/<wp:wrapSquare\b[^>]*>/);
    expect(wrapMatch).not.toBeNull();
    expect(wrapMatch![0]).toContain('wrapText="bothSides"');
  });

  it('defaults wrapText to bothSides for tight wrap', async () => {
    const docXml = await saveFloatingImageXml((image) => image.setWrap('tight'));

    const wrapMatch = docXml.match(/<wp:wrapTight\b[^>]*>/);
    expect(wrapMatch).not.toBeNull();
    expect(wrapMatch![0]).toContain('wrapText="bothSides"');
  });

  it('defaults wrapText to bothSides for through wrap', async () => {
    const docXml = await saveFloatingImageXml((image) => image.setWrap('through'));

    const wrapMatch = docXml.match(/<wp:wrapThrough\b[^>]*>/);
    expect(wrapMatch).not.toBeNull();
    expect(wrapMatch![0]).toContain('wrapText="bothSides"');
  });

  it('keeps an explicitly provided side', async () => {
    const docXml = await saveFloatingImageXml((image) => image.setWrap('square', 'left'));

    const wrapMatch = docXml.match(/<wp:wrapSquare\b[^>]*>/);
    expect(wrapMatch).not.toBeNull();
    expect(wrapMatch![0]).toContain('wrapText="left"');
  });

  it('does not add wrapText to topAndBottom wrap', async () => {
    const docXml = await saveFloatingImageXml((image) => image.setWrap('topAndBottom'));

    const wrapMatch = docXml.match(/<wp:wrapTopAndBottom\b[^>]*\/?>/);
    expect(wrapMatch).not.toBeNull();
    expect(wrapMatch![0]).not.toContain('wrapText');
  });
});
