/**
 * Tracked removeRow/removeRows/removeParagraph must REPLACE the original
 * runs with the w:del-wrapped copies, not append the revision alongside
 * them. Appending serializes the text twice (once live, once as w:delText),
 * so Word shows it duplicated, accepting the revision leaves the "deleted"
 * text alive, and rejecting makes both copies live.
 */
import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function getDocumentXml(doc: Document): Promise<string> {
  const buffer = await doc.toBuffer();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  return zip.getFileAsString('word/document.xml')!;
}

describe('Tracked table deletions replace runs instead of duplicating them', () => {
  let doc: Document;

  afterEach(() => {
    doc.dispose();
  });

  it('removeRow serializes the cell text exactly once, inside w:del', async () => {
    doc = Document.create();
    const table = new Table(2, 1);
    doc.addTable(table);
    table.getRows()[1]!.getCells()[0]!.createParagraph('RowText');

    doc.enableTrackChanges({ author: 'TestAuthor' });
    expect(table.removeRow(1)).toBe(true);

    const para = table.getRows()[1]!.getCells()[0]!.getParagraphs()[0]!;
    const content = para.getContent();
    expect(content.filter((item) => item instanceof Run).length).toBe(0);
    const revisions = content.filter((item) => item instanceof Revision) as Revision[];
    expect(revisions.length).toBe(1);
    expect(revisions[0]!.getType()).toBe('delete');

    const xml = await getDocumentXml(doc);
    expect((xml.match(/RowText/g) || []).length).toBe(1);
    expect(xml).toMatch(/<w:delText[^>]*>RowText/);
    expect(xml).not.toMatch(/<w:t[ >][^<]*RowText/);
  });

  it('removeRows serializes each removed row text exactly once, inside w:del', async () => {
    doc = Document.create();
    const table = new Table(4, 1);
    doc.addTable(table);
    table.getRows()[1]!.getCells()[0]!.createParagraph('SecondRow');
    table.getRows()[2]!.getCells()[0]!.createParagraph('ThirdRow');

    doc.enableTrackChanges({ author: 'TestAuthor' });
    expect(table.removeRows(1, 2)).toBe(true);

    const xml = await getDocumentXml(doc);
    for (const text of ['SecondRow', 'ThirdRow']) {
      expect((xml.match(new RegExp(text, 'g')) || []).length).toBe(1);
      expect(xml).toMatch(new RegExp(`<w:delText[^>]*>${text}`));
      expect(xml).not.toMatch(new RegExp(`<w:t[ >][^<]*${text}`));
    }
  });

  it('TableCell.removeParagraph replaces the runs with the delete revision', async () => {
    doc = Document.create();
    const table = new Table(1, 1);
    doc.addTable(table);
    const cell = table.getRows()[0]!.getCells()[0]!;
    const para = new Paragraph();
    para.addText('CellBody');
    cell.addParagraph(para);

    doc.enableTrackChanges({ author: 'TestAuthor' });
    const index = cell.getParagraphs().indexOf(para);
    expect(cell.removeParagraph(index)).toBe(true);

    const content = para.getContent();
    expect(content.filter((item) => item instanceof Run).length).toBe(0);
    const revisions = content.filter((item) => item instanceof Revision) as Revision[];
    expect(revisions.length).toBe(1);
    expect(revisions[0]!.getType()).toBe('delete');

    const xml = await getDocumentXml(doc);
    expect((xml.match(/CellBody/g) || []).length).toBe(1);
    expect(xml).toMatch(/<w:delText[^>]*>CellBody/);
    expect(xml).not.toMatch(/<w:t[ >][^<]*CellBody/);
  });

  it('accepting the tracked removeRow does not resurrect the deleted text', async () => {
    doc = Document.create();
    const table = new Table(2, 1);
    doc.addTable(table);
    table.getRows()[1]!.getCells()[0]!.createParagraph('Doomed');

    doc.enableTrackChanges({ author: 'TestAuthor' });
    table.removeRow(1);
    await doc.acceptAllRevisions();

    const xml = await getDocumentXml(doc);
    expect(xml).not.toContain('Doomed');
  });
});
