/**
 * Accepting a tracked paragraph-mark deletion (w:del inside w:pPr/w:rPr,
 * ECMA-376 §17.13.5.15) must merge the paragraph's content into the
 * FOLLOWING paragraph and remove the now-empty paragraph — not just clear
 * the marker, which leaves a leftover blank line that Word's own
 * Accept All removes.
 */

import { Document } from '../../src/core/Document';
import { Run } from '../../src/elements/Run';
import { acceptRevisionsInMemory } from '../../src/processors/InMemoryRevisionAcceptor';
import { SelectiveRevisionAcceptor } from '../../src/processors/SelectiveRevisionAcceptor';

describe('C11: accepted paragraph-mark deletion merges with following paragraph', () => {
  let doc: Document;

  afterEach(() => {
    doc.dispose();
  });

  it('merges a body paragraph into its following sibling', () => {
    doc = Document.create();
    const first = doc.createParagraph();
    first.addRun(new Run('First half'));
    first.markParagraphMarkAsDeleted(1, 'Reviewer', new Date('2024-01-01'));
    const second = doc.createParagraph();
    second.addRun(new Run(' second half'));

    const result = acceptRevisionsInMemory(doc);

    expect(result.deletionsAccepted).toBe(1);
    const paragraphs = doc.getAllParagraphs();
    expect(paragraphs.length).toBe(1);
    expect(paragraphs[0]?.getText()).toBe('First half second half');
    expect(paragraphs[0]?.isParagraphMarkDeleted()).toBe(false);
  });

  it('collapses consecutive deleted paragraph marks into one surviving paragraph', () => {
    doc = Document.create();
    const a = doc.createParagraph();
    a.addRun(new Run('A'));
    a.markParagraphMarkAsDeleted(1, 'Reviewer', new Date('2024-01-01'));
    const b = doc.createParagraph();
    b.addRun(new Run('B'));
    b.markParagraphMarkAsDeleted(2, 'Reviewer', new Date('2024-01-01'));
    const c = doc.createParagraph();
    c.addRun(new Run('C'));

    acceptRevisionsInMemory(doc);

    const paragraphs = doc.getAllParagraphs();
    expect(paragraphs.length).toBe(1);
    expect(paragraphs[0]?.getText()).toBe('ABC');
  });

  it('leaves the paragraph in place when it has no following sibling paragraph', () => {
    doc = Document.create();
    const only = doc.createParagraph();
    only.addRun(new Run('Last paragraph'));
    only.markParagraphMarkAsDeleted(1, 'Reviewer', new Date('2024-01-01'));

    const result = acceptRevisionsInMemory(doc);

    expect(result.deletionsAccepted).toBe(1);
    expect(only.isParagraphMarkDeleted()).toBe(false);
    expect(doc.getAllParagraphs().length).toBe(1);
    expect(only.getText()).toBe('Last paragraph');
  });

  it('does not merge across a table sitting between the paragraphs', () => {
    doc = Document.create();
    const before = doc.createParagraph();
    before.addRun(new Run('Before table'));
    before.markParagraphMarkAsDeleted(1, 'Reviewer', new Date('2024-01-01'));
    doc.createTable(1, 1);
    const after = doc.createParagraph();
    after.addRun(new Run('After table'));

    acceptRevisionsInMemory(doc, { cleanupEmptyTables: false });

    expect(before.isParagraphMarkDeleted()).toBe(false);
    expect(before.getText()).toBe('Before table');
    expect(after.getText()).toBe('After table');
    expect(doc.getTables().length).toBe(1);
  });

  it('merges within a table cell so the trailing blank merge target absorbs the content', () => {
    doc = Document.create();
    const table = doc.createTable(1, 1);
    const cell = table.getRows()[0]!.getCells()[0]!;
    const para = cell.createParagraph();
    para.addRun(new Run('Cell text'));
    para.markParagraphMarkAsDeleted(1, 'Reviewer', new Date('2024-01-01'));
    cell.createParagraph();

    acceptRevisionsInMemory(doc);

    const cellParagraphs = cell.getParagraphs();
    expect(cellParagraphs.length).toBe(1);
    expect(cellParagraphs[0]?.getText()).toBe('Cell text');
    expect(cellParagraphs[0]?.isParagraphMarkDeleted()).toBe(false);
  });

  it('SelectiveRevisionAcceptor.accept merges matching paragraph-mark deletions', () => {
    doc = Document.create();
    const first = doc.createParagraph();
    first.addRun(new Run('Selectively'));
    first.markParagraphMarkAsDeleted(3, 'Alice', new Date('2024-02-02'));
    const second = doc.createParagraph();
    second.addRun(new Run(' accepted'));

    const result = SelectiveRevisionAcceptor.accept(doc, { authors: ['Alice'] });

    expect(result.accepted).toContain('3');
    const paragraphs = doc.getAllParagraphs();
    expect(paragraphs.length).toBe(1);
    expect(paragraphs[0]?.getText()).toBe('Selectively accepted');
  });

  it('SelectiveRevisionAcceptor.reject restores the paragraph mark without merging', () => {
    doc = Document.create();
    const first = doc.createParagraph();
    first.addRun(new Run('Kept'));
    first.markParagraphMarkAsDeleted(3, 'Alice', new Date('2024-02-02'));
    const second = doc.createParagraph();
    second.addRun(new Run('Separate'));

    const result = SelectiveRevisionAcceptor.reject(doc, { authors: ['Alice'] });

    expect(result.rejected).toContain('3');
    expect(first.isParagraphMarkDeleted()).toBe(false);
    const paragraphs = doc.getAllParagraphs();
    expect(paragraphs.length).toBe(2);
    expect(paragraphs[0]?.getText()).toBe('Kept');
    expect(paragraphs[1]?.getText()).toBe('Separate');
  });

  it('round-trips the merged result through save and reload', async () => {
    doc = Document.create();
    const first = doc.createParagraph();
    first.addRun(new Run('Round'));
    first.markParagraphMarkAsDeleted(7, 'Reviewer', new Date('2024-01-01'));
    const second = doc.createParagraph();
    second.addRun(new Run('Trip'));

    await doc.acceptAllRevisions();
    const buffer = await doc.toBuffer();
    const reloaded = await Document.loadFromBuffer(buffer);
    try {
      const paragraphs = reloaded.getAllParagraphs();
      expect(paragraphs.length).toBe(1);
      expect(paragraphs[0]?.getText()).toBe('RoundTrip');
    } finally {
      reloaded.dispose();
    }
  });
});
