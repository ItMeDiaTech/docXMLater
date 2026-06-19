/**
 * Tracked addParagraph/addParagraphAt must REPLACE the original runs with
 * the w:ins-wrapped copies, not append the revision alongside them.
 * Appending serializes the text twice (once live, once inside w:ins); the
 * live copy belongs to no revision, so neither accepting nor rejecting the
 * change can ever remove it. Runs already inside an existing revision must
 * not be re-wrapped into a second insertion.
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

describe('Tracked paragraph insertion replaces runs instead of duplicating them', () => {
  let doc: Document;

  afterEach(() => {
    doc.dispose();
  });

  it('addParagraph serializes the text exactly once, inside w:ins', async () => {
    doc = Document.create();
    doc.enableTrackChanges({ author: 'TestAuthor' });

    const para = new Paragraph();
    para.addText('TRACKEDTEXT');
    doc.addParagraph(para);

    const content = para.getContent();
    expect(content.filter((item) => item instanceof Run).length).toBe(0);
    const revisions = content.filter((item) => item instanceof Revision) as Revision[];
    expect(revisions.length).toBe(1);
    expect(revisions[0]!.getType()).toBe('insert');

    const xml = await getDocumentXml(doc);
    expect((xml.match(/TRACKEDTEXT/g) || []).length).toBe(1);
    expect(xml).toMatch(/<w:ins[^>]*>(?:(?!<\/w:ins>)[\s\S])*TRACKEDTEXT/);
  });

  it('accepting the tracked insertion keeps the text exactly once', async () => {
    doc = Document.create();
    doc.enableTrackChanges({ author: 'TestAuthor' });

    const para = new Paragraph();
    para.addText('TRACKEDTEXT');
    doc.addParagraph(para);

    doc.disableTrackChanges();
    await doc.acceptAllRevisions();

    const xml = await getDocumentXml(doc);
    expect((xml.match(/TRACKEDTEXT/g) || []).length).toBe(1);
  });

  it('does not re-wrap runs that are already inside an insertion revision', async () => {
    doc = Document.create();
    doc.enableTrackChanges({ author: 'TestAuthor' });

    const para = new Paragraph();
    para.addRevision(Revision.createInsertion('OrigAuthor', new Run('ALREADYTRACKED')));
    doc.addParagraph(para);

    const revisions = para.getContent().filter((item) => item instanceof Revision) as Revision[];
    expect(revisions.length).toBe(1);
    expect(revisions[0]!.getAuthor()).toBe('OrigAuthor');

    const xml = await getDocumentXml(doc);
    expect((xml.match(/ALREADYTRACKED/g) || []).length).toBe(1);
    expect((xml.match(/<w:ins[\s>]/g) || []).length).toBe(1);
  });

  it('TableCell.addParagraphAt serializes the cell text exactly once, inside w:ins', async () => {
    doc = Document.create();
    doc.enableTrackChanges({ author: 'TestAuthor' });
    const table = new Table(1, 1);
    doc.addTable(table);
    const cell = table.getRows()[0]!.getCells()[0]!;

    const para = new Paragraph();
    para.addText('CELLTEXT');
    cell.addParagraphAt(0, para);

    const content = para.getContent();
    expect(content.filter((item) => item instanceof Run).length).toBe(0);
    const revisions = content.filter((item) => item instanceof Revision) as Revision[];
    expect(revisions.length).toBe(1);
    expect(revisions[0]!.getType()).toBe('insert');

    const xml = await getDocumentXml(doc);
    expect((xml.match(/CELLTEXT/g) || []).length).toBe(1);
    expect(xml).toMatch(/<w:ins[^>]*>(?:(?!<\/w:ins>)[\s\S])*CELLTEXT/);
  });
});
