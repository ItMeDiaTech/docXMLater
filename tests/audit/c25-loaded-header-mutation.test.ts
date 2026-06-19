/**
 * Content added to a loaded header/footer (addText/addParagraph/addTable)
 * must survive save. Previously toXML() unconditionally returned the raw XML
 * preserved at load time, and no mutator invalidated it, so post-load edits
 * vanished silently while the API appeared to succeed.
 *
 * Untouched loaded headers/footers must still round-trip via their preserved
 * raw XML (mutation, not parsing, is what invalidates it).
 */
import { Document } from '../../src/core/Document';
import { Header } from '../../src/elements/Header';
import { Footer } from '../../src/elements/Footer';

const JSZip = require('jszip');

async function buildDocWithHeaderFooter(): Promise<Buffer> {
  const doc = Document.create();
  const header = Header.createDefault();
  header.createParagraph('Original Header');
  const footer = Footer.createDefault();
  footer.createParagraph('Original Footer');
  doc.setHeader(header);
  doc.setFooter(footer);
  doc.createParagraph('Body');
  try {
    return await doc.toBuffer();
  } finally {
    doc.dispose();
  }
}

describe('C25: mutations to loaded headers/footers are saved', () => {
  it('saves text added to a loaded header and footer', async () => {
    const buffer1 = await buildDocWithHeaderFooter();

    const doc = await Document.loadFromBuffer(buffer1);
    let buffer2: Buffer;
    try {
      const headerEntry = doc.getHeaderFooterManager().getAllHeaders()[0]!;
      const footerEntry = doc.getHeaderFooterManager().getAllFooters()[0]!;
      headerEntry.header.addText('DRAFT WATERMARK TEXT');
      footerEntry.footer.addText('ADDED FOOTER TEXT');
      buffer2 = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = await JSZip.loadAsync(buffer2);
    const headerXml = await zip.file('word/header1.xml')!.async('string');
    expect(headerXml).toContain('Original Header');
    expect(headerXml).toContain('DRAFT WATERMARK TEXT');

    const footerXml = await zip.file('word/footer1.xml')!.async('string');
    expect(footerXml).toContain('Original Footer');
    expect(footerXml).toContain('ADDED FOOTER TEXT');
  });

  it('saves a paragraph and a table added to a loaded header', async () => {
    const buffer1 = await buildDocWithHeaderFooter();

    const doc = await Document.loadFromBuffer(buffer1);
    let buffer2: Buffer;
    try {
      const headerEntry = doc.getHeaderFooterManager().getAllHeaders()[0]!;
      headerEntry.header.createParagraph('Second Header Paragraph');
      const table = headerEntry.header.createTable(1, 1);
      table.getCell(0, 0)!.createParagraph('Header Cell Text');
      buffer2 = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = await JSZip.loadAsync(buffer2);
    const headerXml = await zip.file('word/header1.xml')!.async('string');
    expect(headerXml).toContain('Original Header');
    expect(headerXml).toContain('Second Header Paragraph');
    expect(headerXml).toContain('Header Cell Text');
  });

  it('keeps preserved raw XML for loaded headers that are not mutated', async () => {
    const buffer1 = await buildDocWithHeaderFooter();

    const doc = await Document.loadFromBuffer(buffer1);
    try {
      // Parsing repopulates elements through the same mutators; that must not
      // count as a user edit, or raw round-trip fidelity would be lost
      const headerEntry = doc.getHeaderFooterManager().getAllHeaders()[0]!;
      expect(headerEntry.header.getRawXML()).toBeDefined();
      const footerEntry = doc.getHeaderFooterManager().getAllFooters()[0]!;
      expect(footerEntry.footer.getRawXML()).toBeDefined();
    } finally {
      doc.dispose();
    }
  });
});
