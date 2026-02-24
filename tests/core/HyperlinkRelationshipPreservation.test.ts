/**
 * Tests for hyperlink relationship preservation during save pipeline.
 *
 * Verifies that hyperlinks in footnotes, endnotes, and raw nested table content
 * retain their relationship IDs through save/load cycles, and that the orphan
 * cleaner still correctly removes truly orphaned relationships.
 */

import { Document } from '../../src/core/Document';
import { Relationship } from '../../src/core/Relationship';
import { Paragraph } from '../../src/elements/Paragraph';
import { Footnote, FootnoteType } from '../../src/elements/Footnote';
import { Endnote, EndnoteType } from '../../src/elements/Endnote';
import { Table } from '../../src/elements/Table';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { CleanupHelper } from '../../src/helpers/CleanupHelper';

const HYPERLINK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

function addHyperlinkRel(doc: Document, id: string, url: string): void {
  doc
    .getRelationshipManager()
    .addRelationship(
      Relationship.create({ id, type: HYPERLINK_REL_TYPE, target: url, targetMode: 'External' })
    );
}

describe('Hyperlink Relationship Preservation', () => {
  it('should preserve hyperlink relationships in footnotes through save', async () => {
    const doc = Document.create();
    doc.createParagraph('Body text');

    // Create a footnote with a hyperlink
    const footnote = new Footnote({ id: 1, type: FootnoteType.Normal });
    const para = new Paragraph();
    const link = para.addHyperlink('https://example.com/footnote-link');
    link.setText('Footnote Link');
    link.setRelationshipId('rId100');
    footnote.addParagraph(para);

    doc.getFootnoteManager().register(footnote);
    addHyperlinkRel(doc, 'rId100', 'https://example.com/footnote-link');

    const buffer = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(buffer);
    // Per OOXML, footnote hyperlink relationships go in footnotes.xml.rels (part-scoped)
    const relsXml = zip.getFileAsString('word/_rels/footnotes.xml.rels');
    expect(relsXml).toBeDefined();
    expect(relsXml).toContain('rId100');
    expect(relsXml).toContain('https://example.com/footnote-link');
  });

  it('should preserve hyperlink relationships in endnotes through save', async () => {
    const doc = Document.create();
    doc.createParagraph('Body text');

    // Create an endnote with a hyperlink
    const endnote = new Endnote({ id: 1, type: EndnoteType.Normal });
    const para = new Paragraph();
    const link = para.addHyperlink('https://example.com/endnote-link');
    link.setText('Endnote Link');
    link.setRelationshipId('rId200');
    endnote.addParagraph(para);

    doc.getEndnoteManager().register(endnote);
    addHyperlinkRel(doc, 'rId200', 'https://example.com/endnote-link');

    const buffer = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(buffer);
    // Per OOXML, endnote hyperlink relationships go in endnotes.xml.rels (part-scoped)
    const relsXml = zip.getFileAsString('word/_rels/endnotes.xml.rels');
    expect(relsXml).toBeDefined();
    expect(relsXml).toContain('rId200');
    expect(relsXml).toContain('https://example.com/endnote-link');
  });

  it('should preserve hyperlink relationships in raw nested table content through save', async () => {
    const doc = Document.create();

    // Create a table with raw nested content containing a hyperlink reference
    const table = new Table(1, 1);
    const cell = table.getCell(0, 0);
    if (cell) {
      cell.addParagraph(new Paragraph().addText('Cell text'));

      const nestedXml =
        `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>` +
        `<w:tr><w:tc><w:p><w:hyperlink r:id="rId300"><w:r><w:t>Nested Link</w:t></w:r></w:hyperlink></w:p></w:tc></w:tr>` +
        `</w:tbl>`;
      cell.addRawNestedContent(1, nestedXml, 'table');
    }
    doc.addTable(table);

    addHyperlinkRel(doc, 'rId300', 'https://example.com/nested-link');

    const buffer = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(buffer);
    const relsXml = zip.getFileAsString('word/_rels/document.xml.rels');
    expect(relsXml).toBeDefined();
    expect(relsXml).toContain('rId300');
    expect(relsXml).toContain('https://example.com/nested-link');
  });

  it('should still remove truly orphaned hyperlink relationships', async () => {
    const doc = Document.create();
    doc.createParagraph('Body text with no hyperlinks');

    addHyperlinkRel(doc, 'rId999', 'https://example.com/orphaned');

    const buffer = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(buffer);
    const relsXml = zip.getFileAsString('word/_rels/document.xml.rels');
    expect(relsXml).toBeDefined();
    expect(relsXml).not.toContain('rId999');
    expect(relsXml).not.toContain('https://example.com/orphaned');
  });

  it('should preserve hyperlinks in both body and footnotes simultaneously', async () => {
    const doc = Document.create();

    // Add a hyperlink in the body
    const bodyPara = new Paragraph();
    const bodyLink = bodyPara.addHyperlink('https://example.com/body-link');
    bodyLink.setText('Body Link');
    bodyLink.setRelationshipId('rId400');
    doc.addParagraph(bodyPara);

    // Add a hyperlink in a footnote
    const footnote = new Footnote({ id: 1, type: FootnoteType.Normal });
    const fnPara = new Paragraph();
    const fnLink = fnPara.addHyperlink('https://example.com/footnote-link');
    fnLink.setText('Footnote Link');
    fnLink.setRelationshipId('rId401');
    footnote.addParagraph(fnPara);

    doc.getFootnoteManager().register(footnote);

    addHyperlinkRel(doc, 'rId400', 'https://example.com/body-link');
    addHyperlinkRel(doc, 'rId401', 'https://example.com/footnote-link');
    addHyperlinkRel(doc, 'rId999', 'https://example.com/orphaned');

    const buffer = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(buffer);
    // Body hyperlink in document.xml.rels
    const docRelsXml = zip.getFileAsString('word/_rels/document.xml.rels');
    expect(docRelsXml).toBeDefined();
    expect(docRelsXml).toContain('rId400');
    expect(docRelsXml).toContain('https://example.com/body-link');
    expect(docRelsXml).not.toContain('rId999');
    expect(docRelsXml).not.toContain('https://example.com/orphaned');
    // Footnote hyperlink in footnotes.xml.rels (part-scoped per OOXML)
    const fnRelsXml = zip.getFileAsString('word/_rels/footnotes.xml.rels');
    expect(fnRelsXml).toBeDefined();
    expect(fnRelsXml).toContain('rId401');
    expect(fnRelsXml).toContain('https://example.com/footnote-link');
  });

  it('should preserve nested table hyperlink relationships through CleanupHelper', async () => {
    const doc = Document.create();

    // Add a body-level hyperlink
    const bodyPara = new Paragraph();
    const bodyLink = bodyPara.addHyperlink('https://example.com/body');
    bodyLink.setText('Body Link');
    bodyLink.setRelationshipId('rId500');
    doc.addParagraph(bodyPara);
    addHyperlinkRel(doc, 'rId500', 'https://example.com/body');

    // Create a table with a nested table containing a hyperlink in raw XML
    const table = new Table(1, 1);
    const cell = table.getCell(0, 0);
    if (cell) {
      cell.addParagraph(new Paragraph().addText('Outer cell'));

      // Nested table with a hyperlink referencing rId501 (simulates parsed nested table)
      const nestedXml =
        `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>` +
        `<w:tr><w:tc><w:p><w:hyperlink r:id="rId501" w:anchor="section1">` +
        `<w:r><w:t>Nested Link</w:t></w:r></w:hyperlink></w:p></w:tc></w:tr></w:tbl>`;
      cell.addRawNestedContent(1, nestedXml, 'table');
    }
    doc.addTable(table);
    addHyperlinkRel(doc, 'rId501', 'https://example.com/nested');

    // Add an orphaned relationship that should be removed
    addHyperlinkRel(doc, 'rId999', 'https://example.com/orphaned');

    // Run CleanupHelper (which calls cleanupRelationships internally)
    const cleanup = new CleanupHelper(doc);
    cleanup.run({ cleanupRelationships: true });

    // Save and verify
    const buffer = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(buffer);
    const relsXml = zip.getFileAsString('word/_rels/document.xml.rels');
    expect(relsXml).toBeDefined();

    // Body hyperlink preserved
    expect(relsXml).toContain('rId500');
    expect(relsXml).toContain('https://example.com/body');

    // Nested table hyperlink preserved (this was the bug — previously removed as orphaned)
    expect(relsXml).toContain('rId501');
    expect(relsXml).toContain('https://example.com/nested');

    // Truly orphaned relationship removed
    expect(relsXml).not.toContain('rId999');
    expect(relsXml).not.toContain('https://example.com/orphaned');
  });
});
