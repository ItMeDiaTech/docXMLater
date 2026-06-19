/**
 * Per ECMA-376 CT_Anchor (wp.xsd), wp:positionH and wp:positionV are required
 * children of wp:anchor (minOccurs=1). A floating shape configured via
 * setAnchor() without setPosition() must still emit both elements, defaulting
 * to relativeFrom="page" with a zero wp:posOffset, mirroring Image.toXML().
 */
import { Shape } from '../../src/elements/Shape';
import { TextBox } from '../../src/elements/TextBox';
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { inchesToEmus } from '../../src/utils/units';

const ANCHOR_OPTIONS = {
  behindDoc: false,
  locked: false,
  layoutInCell: true,
  allowOverlap: false,
  relativeHeight: 251658240,
};

describe('Floating Shape/TextBox anchor emits required wp:positionH/wp:positionV', () => {
  it('defaults positionH/positionV to page-relative zero offset when no position is set', () => {
    const shape = Shape.createRectangle(inchesToEmus(2), inchesToEmus(1));
    shape.setAnchor(ANCHOR_OPTIONS);
    expect(shape.isFloating()).toBe(true);

    const xml = XMLBuilder.elementToString(shape.toXML());
    expect(xml).toContain('<wp:anchor');
    expect(xml).toContain(
      '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
    );
    expect(xml).toContain(
      '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
    );
  });

  it('orders anchor children per CT_Anchor: simplePos, positionH, positionV, extent', () => {
    const shape = Shape.createRectangle(inchesToEmus(2), inchesToEmus(1));
    shape.setAnchor(ANCHOR_OPTIONS);

    const xml = XMLBuilder.elementToString(shape.toXML());
    const simplePosIdx = xml.indexOf('<wp:simplePos');
    const posHIdx = xml.indexOf('<wp:positionH');
    const posVIdx = xml.indexOf('<wp:positionV');
    const extentIdx = xml.indexOf('<wp:extent');

    expect(simplePosIdx).toBeGreaterThan(-1);
    expect(posHIdx).toBeGreaterThan(simplePosIdx);
    expect(posVIdx).toBeGreaterThan(posHIdx);
    expect(extentIdx).toBeGreaterThan(posVIdx);
  });

  it('keeps explicit position values when setPosition() is used', () => {
    const shape = Shape.createRectangle(inchesToEmus(2), inchesToEmus(1));
    shape.setAnchor(ANCHOR_OPTIONS);
    shape.setPosition(
      { anchor: 'margin', offset: 914400 },
      { anchor: 'paragraph', alignment: 'top' }
    );

    const xml = XMLBuilder.elementToString(shape.toXML());
    expect(xml).toContain(
      '<wp:positionH relativeFrom="margin"><wp:posOffset>914400</wp:posOffset></wp:positionH>'
    );
    expect(xml).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV>'
    );
  });

  it('emits simplePos and default positionH/positionV for anchor-only text boxes', () => {
    const textbox = TextBox.create(inchesToEmus(3), inchesToEmus(2));
    textbox.setAnchor(ANCHOR_OPTIONS);
    expect(textbox.isFloating()).toBe(true);

    const xml = XMLBuilder.elementToString(textbox.toXML());
    expect(xml).toContain('<wp:anchor');
    expect(xml).toContain('<wp:simplePos x="0" y="0"/>');
    expect(xml).toContain(
      '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
    );
    expect(xml).toContain(
      '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>'
    );

    const simplePosIdx = xml.indexOf('<wp:simplePos');
    const posHIdx = xml.indexOf('<wp:positionH');
    const posVIdx = xml.indexOf('<wp:positionV');
    const extentIdx = xml.indexOf('<wp:extent');
    expect(posHIdx).toBeGreaterThan(simplePosIdx);
    expect(posVIdx).toBeGreaterThan(posHIdx);
    expect(extentIdx).toBeGreaterThan(posVIdx);
  });

  it('keeps explicit position values for positioned text boxes', () => {
    const textbox = TextBox.create(inchesToEmus(3), inchesToEmus(2));
    textbox.setAnchor(ANCHOR_OPTIONS);
    textbox.setPosition(
      { anchor: 'margin', offset: 914400 },
      { anchor: 'paragraph', alignment: 'top' }
    );

    const xml = XMLBuilder.elementToString(textbox.toXML());
    expect(xml).toContain('<wp:simplePos x="0" y="0"/>');
    expect(xml).toContain(
      '<wp:positionH relativeFrom="margin"><wp:posOffset>914400</wp:posOffset></wp:positionH>'
    );
    expect(xml).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV>'
    );
  });

  it('saves anchor-only floating shapes with positionH/positionV in document.xml', async () => {
    const doc = Document.create();
    try {
      const para = doc.createParagraph();
      const shape = Shape.createRectangle(inchesToEmus(2), inchesToEmus(1));
      shape.setFill('FF0000');
      shape.setAnchor(ANCHOR_OPTIONS);
      para.addShape(shape);

      const zip = new ZipHandler();
      await zip.loadFromBuffer(await doc.toBuffer());
      const docXml = zip.getFileAsString('word/document.xml')!;

      const anchorMatch = docXml.match(/<wp:anchor[\s\S]*?<\/wp:anchor>/);
      expect(anchorMatch).not.toBeNull();
      expect(anchorMatch![0]).toContain('<wp:positionH relativeFrom="page">');
      expect(anchorMatch![0]).toContain('<wp:positionV relativeFrom="page">');
    } finally {
      doc.dispose();
    }
  });
});
