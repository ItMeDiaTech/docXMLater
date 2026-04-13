/**
 * Tests for theme font stripping in applyStyles().
 *
 * Bug: When a loaded document's style contains theme font attributes
 * (fontAsciiTheme, fontHAnsiTheme, etc.), those attributes survive the
 * spread merge in applyStyles() even when the user provides an explicit font.
 * In OOXML, Word prioritizes theme fonts over explicit font names, so the
 * user's chosen font is silently ignored.
 *
 * Fix: Strip theme font attributes from the merged run config when an
 * explicit font is provided.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

function makeStylesXml(opts: {
  listParagraphThemeFonts?: boolean;
  normalThemeFonts?: boolean;
}): string {
  const listParaRFonts = opts.listParagraphThemeFonts
    ? `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>`
    : `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>`;

  const normalRFonts = opts.normalThemeFonts
    ? `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>`
    : `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr>${normalRFonts}<w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>${listParaRFonts}<w:sz w:val="22"/></w:rPr>
    <w:pPr><w:ind w:left="720"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:b/><w:sz w:val="24"/></w:rPr>
  </w:style>
</w:styles>`;
}

function makeDocumentXml(styleId: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>
      <w:r><w:t>Test content</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
}

async function createDocx(stylesXml: string, documentXml: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();
  zipHandler.addFile('[Content_Types].xml', CONTENT_TYPES_XML);
  zipHandler.addFile('_rels/.rels', RELS_XML);
  zipHandler.addFile('word/_rels/document.xml.rels', DOC_RELS_XML);
  zipHandler.addFile('word/styles.xml', stylesXml);
  zipHandler.addFile('word/document.xml', documentXml);
  return await zipHandler.toBuffer();
}

describe('applyStyles theme font stripping', () => {
  describe('strips theme fonts when explicit font is provided', () => {
    it('should strip theme fonts from ListParagraph when user provides explicit font', async () => {
      const buffer = await createDocx(
        makeStylesXml({ listParagraphThemeFonts: true }),
        makeDocumentXml('ListParagraph')
      );

      const doc = await Document.loadFromBuffer(buffer);
      doc.applyStyles({
        listParagraph: { run: { font: 'Verdana' } },
      });

      const style = (doc as any).stylesManager.getStyle('ListParagraph');
      const runFmt = style?.getRunFormatting();

      expect(runFmt).toBeDefined();
      expect(runFmt.font).toBe('Verdana');
      expect(runFmt.fontAsciiTheme).toBeUndefined();
      expect(runFmt.fontHAnsiTheme).toBeUndefined();
      expect(runFmt.fontEastAsiaTheme).toBeUndefined();
      expect(runFmt.fontCsTheme).toBeUndefined();

      doc.dispose();
    });

    it('should strip theme fonts from Normal when user provides explicit font', async () => {
      const buffer = await createDocx(
        makeStylesXml({ normalThemeFonts: true }),
        makeDocumentXml('Normal')
      );

      const doc = await Document.loadFromBuffer(buffer);
      doc.applyStyles({
        normal: { run: { font: 'Verdana' } },
      });

      const style = (doc as any).stylesManager.getStyle('Normal');
      const runFmt = style?.getRunFormatting();

      expect(runFmt).toBeDefined();
      expect(runFmt.font).toBe('Verdana');
      expect(runFmt.fontAsciiTheme).toBeUndefined();
      expect(runFmt.fontHAnsiTheme).toBeUndefined();

      doc.dispose();
    });

    it('should strip theme fonts from Heading1 when user provides explicit font', async () => {
      const buffer = await createDocx(makeStylesXml({}), makeDocumentXml('Heading1'));

      const doc = await Document.loadFromBuffer(buffer);
      doc.applyStyles({
        heading1: { run: { font: 'Georgia' } },
      });

      const style = (doc as any).stylesManager.getStyle('Heading1');
      const runFmt = style?.getRunFormatting();

      expect(runFmt).toBeDefined();
      expect(runFmt.font).toBe('Georgia');
      expect(runFmt.fontAsciiTheme).toBeUndefined();
      expect(runFmt.fontHAnsiTheme).toBeUndefined();

      doc.dispose();
    });

    it('should strip theme fonts from Heading2 when user provides explicit font', async () => {
      const buffer = await createDocx(makeStylesXml({}), makeDocumentXml('Heading2'));

      const doc = await Document.loadFromBuffer(buffer);
      doc.applyStyles({
        heading2: { run: { font: 'Georgia' } },
      });

      const style = (doc as any).stylesManager.getStyle('Heading2');
      const runFmt = style?.getRunFormatting();

      expect(runFmt).toBeDefined();
      expect(runFmt.font).toBe('Georgia');
      expect(runFmt.fontAsciiTheme).toBeUndefined();
      expect(runFmt.fontHAnsiTheme).toBeUndefined();

      doc.dispose();
    });

    it('should strip theme fonts from Heading3 when user provides explicit font', async () => {
      const buffer = await createDocx(makeStylesXml({}), makeDocumentXml('Heading3'));

      const doc = await Document.loadFromBuffer(buffer);
      doc.applyStyles({
        heading3: { run: { font: 'Georgia' } },
      });

      const style = (doc as any).stylesManager.getStyle('Heading3');
      const runFmt = style?.getRunFormatting();

      expect(runFmt).toBeDefined();
      expect(runFmt.font).toBe('Georgia');
      expect(runFmt.fontAsciiTheme).toBeUndefined();
      expect(runFmt.fontHAnsiTheme).toBeUndefined();

      doc.dispose();
    });
  });

  describe('preserves theme fonts when no explicit font is provided', () => {
    it('should preserve theme fonts on ListParagraph when no font override given', async () => {
      // Create a style with ONLY theme font attributes (no w:ascii/w:hAnsi)
      const themeOnlyStylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr><w:sz w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/><w:sz w:val="22"/></w:rPr>
    <w:pPr><w:ind w:left="720"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
  </w:style>
</w:styles>`;

      const buffer = await createDocx(themeOnlyStylesXml, makeDocumentXml('ListParagraph'));

      const doc = await Document.loadFromBuffer(buffer);
      // Apply styles without specifying a font — theme fonts should survive
      doc.applyStyles({
        listParagraph: { run: { bold: true } },
      });

      const style = (doc as any).stylesManager.getStyle('ListParagraph');
      const runFmt = style?.getRunFormatting();

      expect(runFmt).toBeDefined();
      expect(runFmt.fontAsciiTheme).toBe('minorHAnsi');
      expect(runFmt.fontHAnsiTheme).toBe('minorHAnsi');
      expect(runFmt.fontEastAsiaTheme).toBe('minorEastAsia');
      expect(runFmt.fontCsTheme).toBe('minorBidi');

      doc.dispose();
    });

    it('should strip theme fonts even from existing style when user provides explicit font', async () => {
      // When the existing style has both explicit font AND theme fonts (common in Word docs),
      // and the user provides their own font, theme fonts must be stripped
      const buffer = await createDocx(
        makeStylesXml({ listParagraphThemeFonts: true }),
        makeDocumentXml('ListParagraph')
      );

      const doc = await Document.loadFromBuffer(buffer);
      // The existing style has font: "Calibri" + theme fonts from XML.
      // User does NOT provide a font, so merged config inherits font: "Calibri".
      // Since font is present, theme fonts should be stripped.
      doc.applyStyles({
        listParagraph: { run: { bold: true } },
      });

      const style = (doc as any).stylesManager.getStyle('ListParagraph');
      const runFmt = style?.getRunFormatting();

      expect(runFmt).toBeDefined();
      // The existing style's explicit font "Calibri" triggers theme font stripping
      expect(runFmt.font).toBeDefined();
      expect(runFmt.fontAsciiTheme).toBeUndefined();
      expect(runFmt.fontHAnsiTheme).toBeUndefined();

      doc.dispose();
    });
  });
});
