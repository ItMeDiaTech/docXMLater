/**
 * Tests that hyperlink relationships referenced only from raw-XML passthrough
 * (clickable image a:hlinkClick in wp:docPr, body-level mc:AlternateContent)
 * survive an unmodified load -> save round-trip. The pre-save orphan cleaner
 * must treat these r:id references as in-use, while still removing truly
 * orphaned hyperlink relationships.
 */

import { Document } from '../../src/core/Document';
import { Image } from '../../src/elements/Image';
import { ZipHandler } from '../../src/zip/ZipHandler';

const HYPERLINK_REL =
  '<Relationship Id="rId900" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/click" TargetMode="External"/>';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** 1x1 transparent PNG for image tests */
function createTestPng(): Buffer {
  return Buffer.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);
}

/** Builds a docx containing one inline image, returns its buffer */
async function createDocxWithInlineImage(): Promise<Buffer> {
  const doc = Document.create();
  const image = await Image.fromBuffer(createTestPng(), 'png', 914400, 914400);
  doc.addImage(image);
  doc.createParagraph('Text after image');
  try {
    return await doc.toBuffer();
  } finally {
    doc.dispose();
  }
}

/** Adds a hyperlink relationship entry to word/_rels/document.xml.rels */
function injectHyperlinkRel(relsXml: string, relEntry: string): string {
  expect(relsXml).toContain('</Relationships>');
  return relsXml.replace('</Relationships>', `${relEntry}</Relationships>`);
}

async function roundTrip(buffer: Buffer): Promise<{ documentXml: string; relsXml: string }> {
  const doc = await Document.loadFromBuffer(buffer);
  let saved: Buffer;
  try {
    saved = await doc.toBuffer();
  } finally {
    doc.dispose();
  }

  const zip = new ZipHandler();
  await zip.loadFromBuffer(saved);
  const documentXml = zip.getFileAsString('word/document.xml')!;
  const relsXml = zip.getFileAsString('word/_rels/document.xml.rels')!;
  return { documentXml, relsXml };
}

describe('Hyperlink relationships referenced from raw-XML passthrough', () => {
  it('preserves the relationship for a clickable image (a:hlinkClick in wp:docPr) on round-trip', async () => {
    const base = await createDocxWithInlineImage();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(base);

    // Make the image clickable: add a:hlinkClick to wp:docPr plus its relationship
    const documentXml = zip.getFileAsString('word/document.xml')!;
    const hlinkClick = `<a:hlinkClick xmlns:a="${A_NS}" r:id="rId900"/>`;
    const withClick = documentXml.includes('</wp:docPr>')
      ? documentXml.replace('</wp:docPr>', `${hlinkClick}</wp:docPr>`)
      : documentXml.replace(/<wp:docPr ([^>]*?)\/>/, `<wp:docPr $1>${hlinkClick}</wp:docPr>`);
    expect(withClick).toContain('rId900'); // fixture sanity: injection happened
    zip.updateFile('word/document.xml', withClick);

    const relsXml = zip.getFileAsString('word/_rels/document.xml.rels')!;
    zip.updateFile('word/_rels/document.xml.rels', injectHyperlinkRel(relsXml, HYPERLINK_REL));

    const result = await roundTrip(await zip.toBuffer());

    // The hlinkClick reference is re-emitted from raw passthrough...
    expect(result.documentXml).toContain('rId900');
    // ...so its relationship must not be removed as an orphan
    expect(result.relsXml).toContain('Id="rId900"');
    expect(result.relsXml).toContain('https://example.com/click');
  });

  it('preserves relationships referenced from body-level mc:AlternateContent on round-trip', async () => {
    const doc = Document.create();
    doc.createParagraph('Before alternate content');
    let base: Buffer;
    try {
      base = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = new ZipHandler();
    await zip.loadFromBuffer(base);

    // Inject a body-level AlternateContent block whose drawing references rId900
    const documentXml = zip.getFileAsString('word/document.xml')!;
    const altContent =
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      '<mc:Choice Requires="wps"><w:p><w:r><w:rPr>' +
      `<a:hlinkClick xmlns:a="${A_NS}" r:id="rId900"/>` +
      '</w:rPr></w:r></w:p></mc:Choice>' +
      '<mc:Fallback><w:p/></mc:Fallback>' +
      '</mc:AlternateContent>';
    const withAlt = documentXml.replace('<w:sectPr', `${altContent}<w:sectPr`);
    expect(withAlt).toContain('rId900'); // fixture sanity: injection happened
    zip.updateFile('word/document.xml', withAlt);

    const relsXml = zip.getFileAsString('word/_rels/document.xml.rels')!;
    zip.updateFile('word/_rels/document.xml.rels', injectHyperlinkRel(relsXml, HYPERLINK_REL));

    const result = await roundTrip(await zip.toBuffer());

    expect(result.documentXml).toContain('rId900');
    expect(result.relsXml).toContain('Id="rId900"');
  });

  it('still removes truly orphaned hyperlink relationships', async () => {
    const base = await createDocxWithInlineImage();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(base);

    // Relationship exists but nothing in the document references rId901
    const orphanRel =
      '<Relationship Id="rId901" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/orphan" TargetMode="External"/>';
    const relsXml = zip.getFileAsString('word/_rels/document.xml.rels')!;
    zip.updateFile('word/_rels/document.xml.rels', injectHyperlinkRel(relsXml, orphanRel));

    const result = await roundTrip(await zip.toBuffer());

    expect(result.relsXml).not.toContain('Id="rId901"');
  });
});
