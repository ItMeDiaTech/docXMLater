/**
 * Programmatically generated header/footer parts must declare the full
 * namespace set used by word/document.xml. Header/footer content can
 * legally contain anything body content can — drawings (wp:, a:, pic:,
 * wps:), w14 attributes, mc:AlternateContent — so a root that declares
 * only xmlns:w and xmlns:r yields namespace-ill-formed XML as soon as a
 * Shape or ImageRun is added, and Word reports the part as damaged.
 */
import { Document } from '../../src/core/Document';
import { Header } from '../../src/elements/Header';
import { Footer } from '../../src/elements/Footer';
import { Shape } from '../../src/elements/Shape';

const JSZip = require('jszip');

const DRAWING_NAMESPACES: Record<string, string> = {
  'xmlns:wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
  'xmlns:pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
  'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  'xmlns:wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
};

/** Extracts the root element start tag (everything up to the first '>'). */
function rootStartTag(xml: string, rootName: string): string {
  const start = xml.indexOf(`<${rootName}`);
  expect(start).toBeGreaterThanOrEqual(0);
  return xml.slice(start, xml.indexOf('>', start) + 1);
}

describe('C60: generated header/footer parts declare the full namespace set', () => {
  it('declares drawing and extension namespaces on the w:hdr root', () => {
    const header = Header.createDefault();
    header.addText('Heading');

    const root = rootStartTag(header.toXML(), 'w:hdr');
    for (const [attr, uri] of Object.entries(DRAWING_NAMESPACES)) {
      expect(root).toContain(`${attr}="${uri}"`);
    }
    expect(root).toMatch(/mc:Ignorable="[^"]*w14[^"]*"/);
  });

  it('declares drawing and extension namespaces on the w:ftr root', () => {
    const footer = Footer.createDefault();
    footer.addText('Footing');

    const root = rootStartTag(footer.toXML(), 'w:ftr');
    for (const [attr, uri] of Object.entries(DRAWING_NAMESPACES)) {
      expect(root).toContain(`${attr}="${uri}"`);
    }
    expect(root).toMatch(/mc:Ignorable="[^"]*w14[^"]*"/);
  });

  it('saves a header containing a drawing with every used prefix declared', async () => {
    const doc = Document.create();
    const header = Header.createDefault();
    const para = header.createParagraph();
    para.addShape(Shape.createRectangle(914400, 457200));
    doc.setHeader(header);
    doc.createParagraph('Body');

    let buffer: Buffer;
    try {
      buffer = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = await JSZip.loadAsync(buffer);
    const headerXml = await zip.file('word/header1.xml')!.async('string');

    // The shape serializes via wp:/a:/wps: prefixes...
    expect(headerXml).toContain('<wp:inline');
    expect(headerXml).toContain('<wps:wsp');

    // ...so the part root must declare them or the XML is namespace-ill-formed
    const root = rootStartTag(headerXml, 'w:hdr');
    expect(root).toContain(`xmlns:wp="${DRAWING_NAMESPACES['xmlns:wp']}"`);
    expect(root).toContain(`xmlns:a="${DRAWING_NAMESPACES['xmlns:a']}"`);
    expect(root).toContain(`xmlns:wps="${DRAWING_NAMESPACES['xmlns:wps']}"`);

    // Every prefix used anywhere in the part must have an xmlns declaration
    const declared = new Set([...root.matchAll(/xmlns:([A-Za-z0-9]+)=/g)].map((m) => m[1]));
    const used = new Set([...headerXml.matchAll(/<([A-Za-z0-9]+):/g)].map((m) => m[1]));
    for (const prefix of used) {
      expect(declared).toContain(prefix);
    }
  });
});
