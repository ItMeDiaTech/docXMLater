/**
 * Repairs for three schema violations that make Word report a document as
 * corrupt and that the framework previously emitted or preserved verbatim:
 *
 *  1. `wp:wrapNone` / `wp:wrapTopAndBottom` carrying a `wrapText` attribute.
 *     `wrapText` is only declared on CT_WrapSquare/Tight/Through; a `side`
 *     defaulted during parse leaked it onto the wrap types that forbid it.
 *  2. `wp14:sizeRelH` / `wp14:sizeRelV` holding a bare-text percentage instead
 *     of the required `wp14:pctWidth` / `wp14:pctHeight` child element.
 *  3. `w:rPr` children out of ECMA-376 CT_RPr order (e.g. `w:b` appended after
 *     `w:color`/`w:sz`) in raw-preserved numbering.xml.
 */
import { Document } from '../../src/core/Document';
import { Image } from '../../src/elements/Image';
import { ZipHandler } from '../../src/zip/ZipHandler';
import {
  normalizeAnchorSizeRel,
  reorderRunPropertyChildren,
} from '../../src/utils/xmlSanitization';

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

describe('wrapText is never emitted on wrap types that forbid it', () => {
  it('drops a defaulted side from wrapNone', async () => {
    // Mirrors the parser default that used to leak side="bothSides" onto a
    // plain <wp:wrapNone/> and round-trip an invalid attribute.
    const docXml = await saveFloatingImageXml((image) => image.setWrap('none', 'bothSides'));

    const wrapMatch = docXml.match(/<wp:wrapNone\b[^>]*\/?>/);
    expect(wrapMatch).not.toBeNull();
    expect(wrapMatch![0]).not.toContain('wrapText');
  });

  it('drops a defaulted side from topAndBottom', async () => {
    const docXml = await saveFloatingImageXml((image) =>
      image.setWrap('topAndBottom', 'bothSides')
    );

    const wrapMatch = docXml.match(/<wp:wrapTopAndBottom\b[^>]*\/?>/);
    expect(wrapMatch).not.toBeNull();
    expect(wrapMatch![0]).not.toContain('wrapText');
  });
});

describe('normalizeAnchorSizeRel', () => {
  it('wraps bare-text sizeRelH/sizeRelV in the required pct child', () => {
    const input =
      '<wp14:sizeRelH relativeFrom="margin">0</wp14:sizeRelH>' +
      '<wp14:sizeRelV relativeFrom="margin">0</wp14:sizeRelV>';
    expect(normalizeAnchorSizeRel(input)).toBe(
      '<wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH>' +
        '<wp14:sizeRelV relativeFrom="margin"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV>'
    );
  });

  it('preserves a non-zero percentage value', () => {
    const input = '<wp14:sizeRelH relativeFrom="page">50000</wp14:sizeRelH>';
    expect(normalizeAnchorSizeRel(input)).toBe(
      '<wp14:sizeRelH relativeFrom="page"><wp14:pctWidth>50000</wp14:pctWidth></wp14:sizeRelH>'
    );
  });

  it('leaves an already-valid nested form untouched (idempotent)', () => {
    const valid =
      '<wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH>';
    expect(normalizeAnchorSizeRel(valid)).toBe(valid);
    expect(normalizeAnchorSizeRel(normalizeAnchorSizeRel(valid))).toBe(valid);
  });

  it('ignores unrelated XML', () => {
    const other = '<wp:effectExtent l="0" t="0" r="0" b="0"/>';
    expect(normalizeAnchorSizeRel(other)).toBe(other);
  });
});

describe('reorderRunPropertyChildren', () => {
  it('moves w:b/w:bCs ahead of w:color/w:sz/w:szCs', () => {
    const input =
      '<w:rPr><w:rFonts w:ascii="Symbol"/><w:color w:val="000000"/>' +
      '<w:sz w:val="24"/><w:szCs w:val="24"/>' +
      '<w:b w:val="0"></w:b><w:bCs w:val="0"></w:bCs></w:rPr>';
    expect(reorderRunPropertyChildren(input)).toBe(
      '<w:rPr><w:rFonts w:ascii="Symbol"/>' +
        '<w:b w:val="0"></w:b><w:bCs w:val="0"></w:bCs>' +
        '<w:color w:val="000000"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
    );
  });

  it('is idempotent for already-ordered run properties', () => {
    const ordered =
      '<w:rPr><w:rFonts w:ascii="Symbol"/><w:b/><w:bCs/>' +
      '<w:color w:val="000000"/><w:sz w:val="24"/></w:rPr>';
    expect(reorderRunPropertyChildren(ordered)).toBe(ordered);
  });

  it('reorders every w:rPr block in the part', () => {
    const input =
      '<w:lvl><w:rPr><w:sz w:val="20"/><w:b/></w:rPr></w:lvl>' +
      '<w:lvl><w:rPr><w:color w:val="FF0000"/><w:i/></w:rPr></w:lvl>';
    expect(reorderRunPropertyChildren(input)).toBe(
      '<w:lvl><w:rPr><w:b/><w:sz w:val="20"/></w:rPr></w:lvl>' +
        '<w:lvl><w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr></w:lvl>'
    );
  });

  it('keeps unknown children in relative order at the end', () => {
    const input = '<w:rPr><w:custom/><w:sz w:val="20"/><w:b/></w:rPr>';
    // w:b and w:sz sort into schema order; the unknown w:custom keeps its
    // original relative position after the modeled children.
    expect(reorderRunPropertyChildren(input)).toBe(
      '<w:rPr><w:b/><w:sz w:val="20"/><w:custom/></w:rPr>'
    );
  });

  it('leaves a self-closing w:rPr alone', () => {
    expect(reorderRunPropertyChildren('<w:rPr/>')).toBe('<w:rPr/>');
  });

  it('returns input unchanged when there is no w:rPr', () => {
    const xml = '<w:numbering><w:abstractNum w:abstractNumId="0"/></w:numbering>';
    expect(reorderRunPropertyChildren(xml)).toBe(xml);
  });
});
