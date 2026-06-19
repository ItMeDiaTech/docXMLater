/**
 * Programmatically built header/footer parts must escape XML special
 * characters in text content. The previous bespoke string renderer
 * concatenated text and attribute values verbatim, so header.addText
 * with '&' or '<' produced a malformed part that Word reports as corrupt.
 */
import { Document } from '../../src/core/Document';
import { Header } from '../../src/elements/Header';
import { Footer } from '../../src/elements/Footer';

const JSZip = require('jszip');

describe('X33: header/footer serialization escapes XML special characters', () => {
  it('escapes ampersands and angle brackets in Header.toXML()', () => {
    const header = Header.createDefault();
    header.addText('Smith & Co');
    header.addText('a < b');

    const xml = header.toXML();
    expect(xml).toContain('Smith &amp; Co');
    expect(xml).toContain('a &lt; b');
    expect(xml).not.toContain('Smith & Co');
    expect(xml).not.toContain('a < b');
  });

  it('escapes ampersands and angle brackets in Footer.toXML()', () => {
    const footer = Footer.createDefault();
    footer.addText('Smith & Co');
    footer.addText('a < b');

    const xml = footer.toXML();
    expect(xml).toContain('Smith &amp; Co');
    expect(xml).toContain('a &lt; b');
    expect(xml).not.toContain('Smith & Co');
    expect(xml).not.toContain('a < b');
  });

  it('writes well-formed escaped header/footer parts into the saved document', async () => {
    const doc = Document.create();
    const header = Header.createDefault();
    header.addText('Smith & Co');
    const footer = Footer.createDefault();
    footer.addText('a < b');
    doc.setHeader(header);
    doc.setFooter(footer);
    doc.createParagraph('Body');

    let buffer: Buffer;
    try {
      buffer = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = await JSZip.loadAsync(buffer);
    const headerXml = await zip.file('word/header1.xml')!.async('string');
    expect(headerXml).toContain('Smith &amp; Co');
    expect(headerXml).not.toContain('Smith & Co');

    const footerXml = await zip.file('word/footer1.xml')!.async('string');
    expect(footerXml).toContain('a &lt; b');
    expect(footerXml).not.toContain('<w:t xml:space="preserve">a < b');
  });
});
