/**
 * Word 2010+ emits every Insert > Text Box / WordArt / wps shape as a
 * run-level mc:AlternateContent (ECMA-376 Part 3 — mc:Choice with the
 * DrawingML shape, mc:Fallback with VML):
 *   <w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>…
 * There is no editing model for these, so the run must be preserved
 * verbatim — otherwise the shape and all its w:txbxContent text are
 * silently deleted on save.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

const TEXTBOX_RUN_PARAGRAPH =
  `<w:p>` +
  `<w:r><w:t>LEAD-</w:t></w:r>` +
  `<w:r>` +
  `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
  `<mc:Choice Requires="wps">` +
  `<w:drawing>` +
  `<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
  `<wp:extent cx="914400" cy="914400"/>` +
  `<wp:docPr id="5" name="Text Box 5"/>` +
  `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
  `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  `<wps:txbx><w:txbxContent><w:p><w:r><w:t>TEXTBOX-TEXT</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
  `</wps:wsp>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing>` +
  `</mc:Choice>` +
  `<mc:Fallback>` +
  `<w:pict>` +
  `<v:rect xmlns:v="urn:schemas-microsoft-com:vml" style="width:72pt;height:72pt">` +
  `<v:textbox><w:txbxContent><w:p><w:r><w:t>TEXTBOX-TEXT</w:t></w:r></w:p></w:txbxContent></v:textbox>` +
  `</v:rect>` +
  `</w:pict>` +
  `</mc:Fallback>` +
  `</mc:AlternateContent>` +
  `</w:r>` +
  `<w:r><w:t>-TRAIL</w:t></w:r>` +
  `</w:p>`;

async function buildDocxWithTextBoxRun(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('Run-level AlternateContent round-trip');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);

  const docXml = zip.getFileAsString('word/document.xml')!;
  const updatedDoc = docXml.includes('<w:sectPr')
    ? docXml.replace('<w:sectPr', `${TEXTBOX_RUN_PARAGRAPH}<w:sectPr`)
    : docXml.replace('</w:body>', `${TEXTBOX_RUN_PARAGRAPH}</w:body>`);
  zip.updateFile('word/document.xml', updatedDoc);

  return zip.toBuffer();
}

describe('Run-level mc:AlternateContent (text box / shape) round-trip', () => {
  it('preserves the mc:AlternateContent wrapper and text-box text on unmodified round-trip', async () => {
    const buf1 = await buildDocxWithTextBoxRun();
    const doc = await Document.loadFromBuffer(buf1);
    try {
      const buf2 = await doc.toBuffer();

      const out = new ZipHandler();
      await out.loadFromBuffer(buf2);
      const docXml = out.getFileAsString('word/document.xml')!;
      expect(docXml).toContain('<mc:AlternateContent');
      expect(docXml).toContain('<mc:Choice Requires="wps">');
      expect(docXml).toContain('<mc:Fallback>');
      expect(docXml).toContain('TEXTBOX-TEXT');
      expect(docXml).toContain('<w:txbxContent>');
    } finally {
      doc.dispose();
    }
  });

  it('keeps both the DrawingML choice and the VML fallback intact', async () => {
    const buf1 = await buildDocxWithTextBoxRun();
    const doc = await Document.loadFromBuffer(buf1);
    try {
      const buf2 = await doc.toBuffer();

      const out = new ZipHandler();
      await out.loadFromBuffer(buf2);
      const docXml = out.getFileAsString('word/document.xml')!;
      expect(docXml).toContain('<wps:wsp');
      expect(docXml).toContain('<w:pict>');
      expect(docXml).toContain('<v:rect');
    } finally {
      doc.dispose();
    }
  });

  it('keeps sibling runs around the shape run', async () => {
    const buf1 = await buildDocxWithTextBoxRun();
    const doc = await Document.loadFromBuffer(buf1);
    try {
      const buf2 = await doc.toBuffer();

      const out = new ZipHandler();
      await out.loadFromBuffer(buf2);
      const docXml = out.getFileAsString('word/document.xml')!;
      const leadPos = docXml.indexOf('LEAD-');
      const acPos = docXml.indexOf('<mc:AlternateContent');
      const trailPos = docXml.indexOf('-TRAIL');
      expect(leadPos).toBeGreaterThan(-1);
      expect(acPos).toBeGreaterThan(leadPos);
      expect(trailPos).toBeGreaterThan(acPos);
    } finally {
      doc.dispose();
    }
  });
});
