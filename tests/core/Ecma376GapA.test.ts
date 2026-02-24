/**
 * ECMA-376 Gap Analysis: Phase A Tests (Quick Wins)
 *
 * Tests cover:
 * A1: Underline color attributes
 * A2: CJK paragraph properties
 * A3: Footnote/endnote references in runs
 * A4: tblLayout parsing
 * A5: Theme font references
 * A6: hMerge (legacy horizontal merge)
 * A7: Document background
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Table } from '../../src/elements/Table';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

/**
 * Helper: Creates a DOCX buffer, then injects custom document.xml content.
 */
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

/** Wraps body XML in a minimal document.xml envelope */
function wrapInDocument(bodyContent: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="w14">
  <w:body>
    ${bodyContent}
  </w:body>
</w:document>`;
}

describe('ECMA-376 Gap Analysis Phase A', () => {
  describe('A1: Underline color attributes', () => {
    it('should set and generate underline color', () => {
      const run = new Run('Hello');
      run.setUnderline('single');
      run.setUnderlineColor('FF0000');
      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toContain('w:color="FF0000"');
      expect(xml).toContain('w:val="single"');
    });

    it('should set and generate underline theme color with tint/shade', () => {
      const run = new Run('Hello');
      run.setUnderline('double');
      run.setUnderlineThemeColor('accent1', 0xbf, 0x40);
      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toContain('w:themeColor="accent1"');
      expect(xml).toContain('w:themeTint="BF"');
      expect(xml).toContain('w:themeShade="40"');
    });

    it('should round-trip underline color through parse', async () => {
      const bodyXml = `<w:p><w:r><w:rPr><w:u w:val="single" w:color="0000FF" w:themeColor="hyperlink"/></w:rPr><w:t>test</w:t></w:r></w:p>`;
      const buffer = await createDocxWithDocumentXml(wrapInDocument(bodyXml));
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const runs = doc.getParagraphs()[0]?.getRuns() ?? [];
      expect(runs.length).toBeGreaterThan(0);
      const fmt = runs[0]!.getFormatting();
      expect(fmt.underlineColor).toBe('0000FF');
      expect(fmt.underlineThemeColor).toBe('hyperlink');
      doc.dispose();
    });
  });

  describe('A2: CJK paragraph properties', () => {
    it('should set and generate kinsoku property', () => {
      const para = new Paragraph();
      para.addText('CJK text');
      para.setKinsoku(true);
      const xml = XMLBuilder.elementToString(para.toXML());
      expect(xml).toContain('<w:kinsoku');
      expect(xml).toContain('w:val="1"');
    });

    it('should set and generate all 6 CJK properties', () => {
      const para = new Paragraph();
      para.addText('CJK');
      para.setKinsoku(false);
      para.setWordWrap(false);
      para.setOverflowPunct(true);
      para.setTopLinePunct(true);
      para.setAutoSpaceDE(false);
      para.setAutoSpaceDN(false);
      const xml = XMLBuilder.elementToString(para.toXML());
      expect(xml).toContain('<w:kinsoku');
      expect(xml).toContain('<w:wordWrap');
      expect(xml).toContain('<w:overflowPunct');
      expect(xml).toContain('<w:topLinePunct');
      expect(xml).toContain('<w:autoSpaceDE');
      expect(xml).toContain('<w:autoSpaceDN');
    });

    it('should round-trip CJK properties through parse', async () => {
      const bodyXml = `<w:p><w:pPr><w:kinsoku w:val="0"/><w:wordWrap w:val="0"/><w:overflowPunct w:val="1"/></w:pPr><w:r><w:t>CJK</w:t></w:r></w:p>`;
      const buffer = await createDocxWithDocumentXml(wrapInDocument(bodyXml));
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const para = doc.getParagraphs()[0]!;
      expect(para).toBeDefined();
      expect(para.getFormatting().kinsoku).toBe(false);
      expect(para.getFormatting().wordWrap).toBe(false);
      expect(para.getFormatting().overflowPunct).toBe(true);
      doc.dispose();
    });
  });

  describe('A3: Footnote/endnote references in runs', () => {
    it('should generate footnoteReference in toXML', () => {
      const run = Run.createFromContent([{ type: 'footnoteReference', footnoteId: 1 }]);
      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toContain('<w:footnoteReference');
      expect(xml).toContain('w:id="1"');
    });

    it('should generate endnoteReference in toXML', () => {
      const run = Run.createFromContent([{ type: 'endnoteReference', endnoteId: 2 }]);
      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toContain('<w:endnoteReference');
      expect(xml).toContain('w:id="2"');
    });

    it('should round-trip footnoteReference through parse', async () => {
      const bodyXml = `<w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="3"/></w:r></w:p>`;
      const buffer = await createDocxWithDocumentXml(wrapInDocument(bodyXml));
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const runs = doc.getParagraphs()[0]?.getRuns() ?? [];
      const content = runs[0]?.getContent() ?? [];
      const fnRef = content.find((c) => c.type === 'footnoteReference');
      expect(fnRef).toBeDefined();
      expect(fnRef!.footnoteId).toBe(3);
      doc.dispose();
    });
  });

  describe('A4: tblLayout parsing', () => {
    it('should round-trip tblLayout type="fixed" through parse', async () => {
      const bodyXml = `<w:tbl>
        <w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
        <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>`;
      const buffer = await createDocxWithDocumentXml(wrapInDocument(bodyXml));
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const tables = doc.getTables();
      expect(tables.length).toBeGreaterThan(0);
      const layout = tables[0]!.getLayout?.();
      expect(layout).toBe('fixed');
      doc.dispose();
    });
  });

  describe('A5: Theme font references', () => {
    it('should set and generate theme font attributes', () => {
      const run = new Run('Theme font');
      run.setFontAsciiTheme('majorHAnsi');
      run.setFontHAnsiTheme('majorHAnsi');
      run.setFontEastAsiaTheme('majorEastAsia');
      run.setFontCsTheme('majorBidi');
      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toContain('w:asciiTheme="majorHAnsi"');
      expect(xml).toContain('w:hAnsiTheme="majorHAnsi"');
      expect(xml).toContain('w:eastAsiaTheme="majorEastAsia"');
      expect(xml).toContain('w:cstheme="majorBidi"');
    });

    it('should round-trip theme fonts through parse', async () => {
      const bodyXml = `<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/></w:rPr><w:t>test</w:t></w:r></w:p>`;
      const buffer = await createDocxWithDocumentXml(wrapInDocument(bodyXml));
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const runs = doc.getParagraphs()[0]?.getRuns() ?? [];
      const fmt = runs[0]!.getFormatting();
      expect(fmt.fontAsciiTheme).toBe('minorHAnsi');
      expect(fmt.fontHAnsiTheme).toBe('minorHAnsi');
      expect(fmt.fontEastAsiaTheme).toBe('minorEastAsia');
      expect(fmt.fontCsTheme).toBe('minorBidi');
      doc.dispose();
    });
  });

  describe('A6: hMerge (legacy horizontal merge)', () => {
    it('should round-trip hMerge through parse', async () => {
      const bodyXml = `<w:tbl>
        <w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:hMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Merged</w:t></w:r></w:p></w:tc>
          <w:tc><w:tcPr><w:hMerge/></w:tcPr><w:p/></w:tc>
        </w:tr>
      </w:tbl>`;
      const buffer = await createDocxWithDocumentXml(wrapInDocument(bodyXml));
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const tables = doc.getTables();
      expect(tables.length).toBeGreaterThan(0);
      const row = tables[0]!.getRow(0)!;
      expect(row).toBeDefined();
      expect(row.getCell(0)!.getHorizontalMerge?.()).toBe('restart');
      expect(row.getCell(1)!.getHorizontalMerge?.()).toBe('continue');
      doc.dispose();
    });

    it('should generate hMerge in toXML', () => {
      const table = new Table(1, 2);
      table.getRow(0)!.getCell(0)!.setHorizontalMerge('restart');
      table.getRow(0)!.getCell(1)!.setHorizontalMerge('continue');
      const xml = XMLBuilder.elementToString(table.toXML());
      expect(xml).toContain('<w:hMerge');
      expect(xml).toContain('w:val="restart"');
    });
  });

  describe('A7: Document background', () => {
    it('should set and get document background', () => {
      const doc = Document.create();
      doc.setDocumentBackground({ color: 'FFFFFF' });
      const bg = doc.getDocumentBackground();
      expect(bg?.color).toBe('FFFFFF');
      doc.dispose();
    });

    it('should round-trip document background through parse', async () => {
      const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:background w:color="E6E6E6" w:themeColor="background1" w:themeShade="E6"/>
  <w:body>
    <w:p><w:r><w:t>Test</w:t></w:r></w:p>
  </w:body>
</w:document>`;
      const buffer = await createDocxWithDocumentXml(documentXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const bg = doc.getDocumentBackground();
      expect(bg).toBeDefined();
      expect(bg!.color).toBe('E6E6E6');
      expect(bg!.themeColor).toBe('background1');
      expect(bg!.themeShade).toBe('E6');
      doc.dispose();
    });

    it('should include background in generated document.xml', async () => {
      const doc = Document.create();
      doc.addParagraph(new Paragraph().addText('test'));
      doc.setDocumentBackground({ color: 'AABBCC' });
      const buffer = await doc.toBuffer();
      doc.dispose();

      const zipHandler = new ZipHandler();
      await zipHandler.loadFromBuffer(buffer);
      const docXml = zipHandler.getFileAsString(DOCX_PATHS.DOCUMENT);
      expect(docXml).toContain('<w:background');
      expect(docXml).toContain('w:color="AABBCC"');
    });
  });
});
