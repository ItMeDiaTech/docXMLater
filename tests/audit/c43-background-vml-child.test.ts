/**
 * Regression tests: w:background with child content must survive the
 * round-trip. Per ECMA-376 §17.2.1 w:background may carry a v:background
 * child holding the actual picture/gradient/texture page fill — Word emits
 * the expanded element form for those, and only flat-color backgrounds are
 * self-closing. A self-closing-only parse drops the entire page background
 * (attributes and fill) because word/document.xml is regenerated on save.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

/** Creates a DOCX buffer, then injects custom document.xml content. */
async function createDocxWithDocumentXml(documentXml: string): Promise<Buffer> {
  const doc = Document.create();
  doc.addParagraph(new Paragraph().addText('placeholder'));
  const buffer = await doc.toBuffer();
  doc.dispose();

  const zipHandler = new ZipHandler();
  await zipHandler.loadFromBuffer(buffer);
  zipHandler.updateFile(DOCX_PATHS.DOCUMENT, documentXml);
  return await zipHandler.toBuffer();
}

const EXPANDED_BACKGROUND_DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:o="urn:schemas-microsoft-com:office:office">
  <w:background w:color="C0504D" w:themeColor="accent2"><v:background id="_x0000_s1025" o:bwmode="white" o:targetscreensize="1024,768"><v:fill color2="silver" focus="100%" type="gradient"/></v:background></w:background>
  <w:body>
    <w:p><w:r><w:t>Test</w:t></w:r></w:p>
  </w:body>
</w:document>`;

describe('w:background with v:background child (C43)', () => {
  it('parses the expanded form: attributes plus rawInnerXml passthrough', async () => {
    const buffer = await createDocxWithDocumentXml(EXPANDED_BACKGROUND_DOCUMENT);
    const doc = await Document.loadFromBuffer(buffer);
    try {
      const bg = doc.getDocumentBackground() as
        | { color?: string; themeColor?: string; rawInnerXml?: string }
        | undefined;
      expect(bg).toBeDefined();
      expect(bg!.color).toBe('C0504D');
      expect(bg!.themeColor).toBe('accent2');
      expect(bg!.rawInnerXml).toContain('<v:background');
      expect(bg!.rawInnerXml).toContain('type="gradient"');
    } finally {
      doc.dispose();
    }
  });

  it('re-emits the expanded form with the VML fill on save', async () => {
    const buffer = await createDocxWithDocumentXml(EXPANDED_BACKGROUND_DOCUMENT);
    const doc = await Document.loadFromBuffer(buffer);
    let outputBuffer: Buffer;
    try {
      outputBuffer = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = new ZipHandler();
    await zip.loadFromBuffer(outputBuffer);
    const docXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT) || '';

    // Expanded element form with attributes and the verbatim v:background child
    expect(docXml).toMatch(
      /<w:background[^>]*w:color="C0504D"[^>]*><v:background[\s\S]*?<\/v:background><\/w:background>/
    );
    expect(docXml).toContain('o:targetscreensize="1024,768"');
    expect(docXml).toContain('<v:fill color2="silver" focus="100%" type="gradient"/>');

    // And the fill survives a second load
    const reloaded = await Document.loadFromBuffer(outputBuffer);
    try {
      const bg = reloaded.getDocumentBackground() as { rawInnerXml?: string } | undefined;
      expect(bg?.rawInnerXml).toContain('<v:background');
    } finally {
      reloaded.dispose();
    }
  });

  it('keeps flat-color backgrounds self-closing', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:background w:color="E6E6E6"/>
  <w:body>
    <w:p><w:r><w:t>Test</w:t></w:r></w:p>
  </w:body>
</w:document>`;
    const buffer = await createDocxWithDocumentXml(documentXml);
    const doc = await Document.loadFromBuffer(buffer);
    let outputBuffer: Buffer;
    try {
      expect(doc.getDocumentBackground()?.color).toBe('E6E6E6');
      outputBuffer = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = new ZipHandler();
    await zip.loadFromBuffer(outputBuffer);
    const docXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT) || '';
    expect(docXml).toMatch(/<w:background[^>]*w:color="E6E6E6"[^>]*\/>/);
    expect(docXml).not.toContain('</w:background>');
  });
});
