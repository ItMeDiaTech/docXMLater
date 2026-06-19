/**
 * Tracked Document.removeParagraph must REPLACE the original runs with the
 * w:del-wrapped copies, not append the revision alongside them. Appending
 * serializes the text twice (live w:t plus w:delText), and accepting all
 * revisions removes only the w:del copy — the "deleted" paragraph survives
 * forever. Per ECMA-376 §17.13.5.14 deleted content must exist only inside
 * w:del. Hyperlink content must be wrapped too, or it can never be deleted
 * under tracking.
 */
import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { Revision } from '../../src/elements/Revision';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function getDocumentXml(doc: Document): Promise<string> {
  const buffer = await doc.toBuffer();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  return zip.getFileAsString('word/document.xml')!;
}

describe('Tracked removeParagraph replaces runs instead of duplicating them', () => {
  let doc: Document;

  afterEach(() => {
    doc.dispose();
  });

  it('serializes the text exactly once, as w:delText inside w:del', async () => {
    doc = Document.create();
    const para = new Paragraph();
    para.addText('DELETEME');
    doc.addParagraph(para);

    doc.enableTrackChanges({ author: 'TestAuthor' });
    expect(doc.removeParagraph(para)).toBe(true);

    // Paragraph stays in the document, marked deleted via revision
    expect(doc.getParagraphs()).toContain(para);
    const content = para.getContent();
    expect(content.filter((item) => item instanceof Run).length).toBe(0);
    const revisions = content.filter((item) => item instanceof Revision) as Revision[];
    expect(revisions.length).toBe(1);
    expect(revisions[0]!.getType()).toBe('delete');

    const xml = await getDocumentXml(doc);
    expect((xml.match(/DELETEME/g) || []).length).toBe(1);
    expect(xml).toMatch(/<w:delText[^>]*>DELETEME/);
    expect(xml).not.toMatch(/<w:t[ >][^<]*DELETEME/);
  });

  it('accepting the tracked removal actually removes the text', async () => {
    doc = Document.create();
    const para = new Paragraph();
    para.addText('DELETEME');
    doc.addParagraph(para);

    doc.enableTrackChanges({ author: 'TestAuthor' });
    doc.removeParagraph(para);
    doc.disableTrackChanges();
    await doc.acceptAllRevisions();

    const xml = await getDocumentXml(doc);
    expect(xml).not.toContain('DELETEME');
  });

  it('wraps top-level hyperlink content in the delete revision', () => {
    doc = Document.create();
    const para = new Paragraph();
    para.addHyperlink(new Hyperlink({ url: 'https://example.com', text: 'LinkText' }));
    doc.addParagraph(para);

    doc.enableTrackChanges({ author: 'TestAuthor' });
    expect(doc.removeParagraph(para)).toBe(true);

    const content = para.getContent();
    expect(content.filter((item) => item instanceof Hyperlink).length).toBe(0);
    const revisions = content.filter((item) => item instanceof Revision) as Revision[];
    expect(revisions.length).toBe(1);
    expect(revisions[0]!.getType()).toBe('delete');
  });

  it('TableCell.removeParagraph wraps hyperlink content in the delete revision', () => {
    doc = Document.create();
    const table = new Table(1, 1);
    doc.addTable(table);
    const cell = table.getRows()[0]!.getCells()[0]!;
    const para = new Paragraph();
    para.addHyperlink(new Hyperlink({ url: 'https://example.com', text: 'CellLink' }));
    cell.addParagraph(para);

    doc.enableTrackChanges({ author: 'TestAuthor' });
    const index = cell.getParagraphs().indexOf(para);
    expect(cell.removeParagraph(index)).toBe(true);

    const content = para.getContent();
    expect(content.filter((item) => item instanceof Hyperlink).length).toBe(0);
    const revisions = content.filter((item) => item instanceof Revision) as Revision[];
    expect(revisions.length).toBe(1);
    expect(revisions[0]!.getType()).toBe('delete');
  });
});
