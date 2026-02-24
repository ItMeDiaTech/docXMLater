/**
 * Integration test for relative TOC indentation via indentPerLevel.
 *
 * Creates a document with a TOC and Heading 2 / Heading 3 content,
 * applies formatTOCStyles() with indentPerLevel, saves to a buffer,
 * and verifies the resulting style indentation both in-memory and in
 * the generated styles.xml.
 */

import { Document, TableOfContents, TableOfContentsElement } from '../../src';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('TOC relative indentation (indentPerLevel) integration', () => {
  /**
   * Helper: build a document with a TOC covering heading levels 1-3,
   * two Heading 2 paragraphs, and two Heading 3 paragraphs.
   */
  function buildDocument(indentPerLevel: number): Document {
    const doc = Document.create();

    // Insert TOC at the top
    const toc = new TableOfContents({ title: 'Contents', levels: 3 });
    doc.insertTocAt(0, TableOfContentsElement.create(toc));

    // Add heading 2 and heading 3 content
    doc.createParagraph('First Section').setStyle('Heading2');
    doc.createParagraph('Body text under first section.');

    doc.createParagraph('First Subsection').setStyle('Heading3');
    doc.createParagraph('Body text under first subsection.');

    doc.createParagraph('Second Section').setStyle('Heading2');
    doc.createParagraph('Body text under second section.');

    doc.createParagraph('Second Subsection').setStyle('Heading3');
    doc.createParagraph('Body text under second subsection.');

    // Format TOC styles with relative indentation — only levels 2 & 3
    doc.formatTOCStyles({
      run: { font: 'Verdana', size: 12 },
      paragraph: { spacing: { before: 0, after: 0 } },
      levels: [2, 3],
      indentPerLevel,
    });

    return doc;
  }

  /**
   * Helper: extract a style block from styles.xml inside a DOCX buffer.
   */
  async function getStyleXmlFromBuffer(
    buffer: Buffer,
    styleId: string
  ): Promise<string | undefined> {
    const zip = new ZipHandler();
    await zip.loadFromBuffer(buffer);
    const stylesXml = zip.getFileAsString('word/styles.xml');
    if (!stylesXml) return undefined;
    const pattern = new RegExp(`<w:style[^>]*\\sw:styleId="${styleId}"[^>]*>[\\s\\S]*?</w:style>`);
    return stylesXml.match(pattern)?.[0];
  }

  // -- In-memory tests --

  it('should set TOC2=0 and TOC3=360 in memory when indentPerLevel is 360', () => {
    const doc = buildDocument(360);

    expect(doc.getStyle('TOC2')?.getParagraphFormatting()?.indentation?.left).toBe(0);
    expect(doc.getStyle('TOC3')?.getParagraphFormatting()?.indentation?.left).toBe(360);

    doc.dispose();
  });

  it('should make all levels flush-left in memory when indentPerLevel is 0', () => {
    const doc = buildDocument(0);

    expect(doc.getStyle('TOC2')?.getParagraphFormatting()?.indentation?.left).toBe(0);
    expect(doc.getStyle('TOC3')?.getParagraphFormatting()?.indentation?.left).toBe(0);

    doc.dispose();
  });

  it('should preserve run formatting alongside relative indent', () => {
    const doc = buildDocument(360);

    expect(doc.getStyle('TOC2')?.getRunFormatting()?.font).toBe('Verdana');
    expect(doc.getStyle('TOC2')?.getRunFormatting()?.size).toBe(12);
    expect(doc.getStyle('TOC3')?.getRunFormatting()?.font).toBe('Verdana');
    expect(doc.getStyle('TOC3')?.getRunFormatting()?.size).toBe(12);

    doc.dispose();
  });

  it('should preserve spacing alongside relative indent', () => {
    const doc = buildDocument(360);

    expect(doc.getStyle('TOC2')?.getParagraphFormatting()?.spacing?.before).toBe(0);
    expect(doc.getStyle('TOC2')?.getParagraphFormatting()?.spacing?.after).toBe(0);
    expect(doc.getStyle('TOC3')?.getParagraphFormatting()?.spacing?.before).toBe(0);
    expect(doc.getStyle('TOC3')?.getParagraphFormatting()?.spacing?.after).toBe(0);

    doc.dispose();
  });

  it('should not create TOC1 style when only levels [2, 3] are requested', () => {
    const doc = buildDocument(360);

    expect(doc.getStyle('TOC1')).toBeUndefined();
    expect(doc.getStyle('TOC4')).toBeUndefined();

    doc.dispose();
  });

  // -- Saved DOCX / XML verification tests --

  it('should write correct w:ind values to styles.xml', async () => {
    const doc = buildDocument(360);
    const buffer = await doc.toBuffer();
    doc.dispose();

    const toc2Xml = await getStyleXmlFromBuffer(buffer, 'TOC2');
    const toc3Xml = await getStyleXmlFromBuffer(buffer, 'TOC3');

    // TOC2 should have w:left="0"
    expect(toc2Xml).toContain('w:left="0"');

    // TOC3 should have w:left="360"
    expect(toc3Xml).toContain('w:left="360"');
  });

  it('should produce a valid DOCX buffer', async () => {
    const doc = buildDocument(360);
    const buffer = await doc.toBuffer();
    doc.dispose();

    // PK ZIP signature
    expect(buffer[0]).toBe(0x50);
    expect(buffer[1]).toBe(0x4b);
    expect(buffer.length).toBeGreaterThan(1000);
  });

  it('should include heading paragraphs in the saved document', async () => {
    const doc = buildDocument(360);
    const buffer = await doc.toBuffer();
    doc.dispose();

    const loaded = await Document.loadFromBuffer(buffer);
    const paragraphs = loaded.getParagraphs();

    const headingTexts = paragraphs
      .filter((p) => {
        const style = p.getStyle?.();
        return style === 'Heading2' || style === 'Heading3';
      })
      .map((p) => p.getText());

    expect(headingTexts).toContain('First Section');
    expect(headingTexts).toContain('Second Section');
    expect(headingTexts).toContain('First Subsection');
    expect(headingTexts).toContain('Second Subsection');

    loaded.dispose();
  });

  it('should only include TOC2 and TOC3 styles in styles.xml (not TOC1)', async () => {
    const doc = buildDocument(360);
    const buffer = await doc.toBuffer();
    doc.dispose();

    const toc1Xml = await getStyleXmlFromBuffer(buffer, 'TOC1');
    const toc2Xml = await getStyleXmlFromBuffer(buffer, 'TOC2');
    const toc3Xml = await getStyleXmlFromBuffer(buffer, 'TOC3');

    expect(toc1Xml).toBeUndefined();
    expect(toc2Xml).toBeDefined();
    expect(toc3Xml).toBeDefined();
  });

  it('should write flat indentation to styles.xml when indentPerLevel is 0', async () => {
    const doc = buildDocument(0);
    const buffer = await doc.toBuffer();
    doc.dispose();

    const toc2Xml = await getStyleXmlFromBuffer(buffer, 'TOC2');
    const toc3Xml = await getStyleXmlFromBuffer(buffer, 'TOC3');

    expect(toc2Xml).toContain('w:left="0"');
    expect(toc3Xml).toContain('w:left="0"');
  });
});
