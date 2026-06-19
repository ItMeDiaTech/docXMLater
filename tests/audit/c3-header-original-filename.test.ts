/**
 * Loaded header/footer parts must keep their original part names. The parser
 * returns the relationship target as `filename`, and registration must store
 * it; otherwise saveHeaders()/saveFooters() write edits to a renumbered orphan
 * part (header1.xml) while the relationship still targets the original part
 * (header2.xml) containing stale content — edits are silently lost.
 */
import { Document } from '../../src/core/Document';
import { Header } from '../../src/elements/Header';
import { Footer } from '../../src/elements/Footer';

const JSZip = require('jszip');

/**
 * Builds a docx whose only header part is word/header2.xml and only footer
 * part is word/footer3.xml — part numbering that does not match the
 * registration order docxmlater would generate.
 */
async function buildDocWithRenamedParts(): Promise<Buffer> {
  const doc = Document.create();
  const header = Header.createDefault();
  header.createParagraph('Original Header Text');
  doc.setHeader(header);
  const footer = Footer.createDefault();
  footer.createParagraph('Original Footer Text');
  doc.setFooter(footer);
  doc.createParagraph('Body');
  const buf = await doc.toBuffer();
  doc.dispose();

  const zip = await JSZip.loadAsync(buf);

  const headerXml = await zip.file('word/header1.xml')!.async('string');
  zip.remove('word/header1.xml');
  zip.file('word/header2.xml', headerXml);

  const footerXml = await zip.file('word/footer1.xml')!.async('string');
  zip.remove('word/footer1.xml');
  zip.file('word/footer3.xml', footerXml);

  let rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
  rels = rels.replace('Target="header1.xml"', 'Target="header2.xml"');
  rels = rels.replace('Target="footer1.xml"', 'Target="footer3.xml"');
  zip.file('word/_rels/document.xml.rels', rels);

  let contentTypes = await zip.file('[Content_Types].xml')!.async('string');
  contentTypes = contentTypes.replace('/word/header1.xml', '/word/header2.xml');
  contentTypes = contentTypes.replace('/word/footer1.xml', '/word/footer3.xml');
  zip.file('[Content_Types].xml', contentTypes);

  return zip.generateAsync({ type: 'nodebuffer' });
}

function getReferenceRId(docXml: string, tag: string, type: string): string {
  const match = new RegExp(`<w:${tag}[^>]*w:type="${type}"[^>]*/>`).exec(docXml);
  expect(match).not.toBeNull();
  const rId = /r:id="([^"]+)"/.exec(match![0]);
  expect(rId).not.toBeNull();
  return rId![1]!;
}

function getRelTarget(relsXml: string, rId: string): string {
  const match = new RegExp(`<Relationship[^>]*Id="${rId}"[^>]*/>`).exec(relsXml);
  expect(match).not.toBeNull();
  const target = /Target="([^"]+)"/.exec(match![0]);
  expect(target).not.toBeNull();
  return target![1]!;
}

describe('C3: loaded headers/footers keep their original part filenames', () => {
  it('registers loaded parts under their relationship targets', async () => {
    const buf = await buildDocWithRenamedParts();
    const doc = await Document.loadFromBuffer(buf);
    try {
      const manager = doc.getHeaderFooterManager();
      expect(manager.getAllHeaders().map((e) => e.filename)).toEqual(['header2.xml']);
      expect(manager.getAllFooters().map((e) => e.filename)).toEqual(['footer3.xml']);
    } finally {
      doc.dispose();
    }
  });

  it('saves header/footer edits into the part the relationship targets', async () => {
    const buf = await buildDocWithRenamedParts();
    const doc = await Document.loadFromBuffer(buf);
    let saved: Buffer;
    try {
      const headerEntry = doc.getHeaderFooterManager().getAllHeaders()[0]!;
      headerEntry.header.clear();
      headerEntry.header.createParagraph('Edited Header Text');

      const footerEntry = doc.getHeaderFooterManager().getAllFooters()[0]!;
      footerEntry.footer.clear();
      footerEntry.footer.createParagraph('Edited Footer Text');

      saved = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = await JSZip.loadAsync(saved);

    // No renumbered orphan parts
    expect(zip.file('word/header1.xml')).toBeNull();
    expect(zip.file('word/footer1.xml')).toBeNull();

    // Edits landed in the original parts
    const headerXml = await zip.file('word/header2.xml')!.async('string');
    expect(headerXml).toContain('Edited Header Text');
    expect(headerXml).not.toContain('Original Header Text');

    const footerXml = await zip.file('word/footer3.xml')!.async('string');
    expect(footerXml).toContain('Edited Footer Text');
    expect(footerXml).not.toContain('Original Footer Text');

    // Relationships referenced from the sectPr still resolve to those parts
    const docXml = await zip.file('word/document.xml')!.async('string');
    const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string');
    const headerRId = getReferenceRId(docXml, 'headerReference', 'default');
    const footerRId = getReferenceRId(docXml, 'footerReference', 'default');
    expect(getRelTarget(relsXml, headerRId)).toBe('header2.xml');
    expect(getRelTarget(relsXml, footerRId)).toBe('footer3.xml');

    // Content types declare the original part names
    const contentTypes = await zip.file('[Content_Types].xml')!.async('string');
    expect(contentTypes).toContain('/word/header2.xml');
    expect(contentTypes).toContain('/word/footer3.xml');
  });

  it('clearAllHeaderFooterContent applies to the referenced parts', async () => {
    const buf = await buildDocWithRenamedParts();
    const doc = await Document.loadFromBuffer(buf);
    let saved: Buffer;
    try {
      doc.clearAllHeaderFooterContent();
      saved = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = await JSZip.loadAsync(saved);
    const headerXml = await zip.file('word/header2.xml')!.async('string');
    expect(headerXml).not.toContain('Original Header Text');
    const footerXml = await zip.file('word/footer3.xml')!.async('string');
    expect(footerXml).not.toContain('Original Footer Text');
  });
});
