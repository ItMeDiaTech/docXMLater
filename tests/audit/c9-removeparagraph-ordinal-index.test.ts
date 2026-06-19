/**
 * Document.removeParagraph(number) — the numeric overload must use
 * paragraph-ordinal indexing, matching getParagraphAt/getParagraphIndex
 * and removeTable(number).
 *
 * It previously indexed straight into the mixed bodyElements array, so
 * whenever a table (or other non-paragraph element) preceded the target,
 * removeParagraph(i) silently deleted the wrong paragraph while still
 * returning true. removeParagraph(doc.getParagraphIndex(p)) therefore
 * corrupted any mixed-content document.
 */

import { Document } from '../../src/core/Document';

describe('removeParagraph(number) uses paragraph-ordinal indexing', () => {
  let doc: Document;

  beforeEach(() => {
    doc = Document.create();
  });

  afterEach(() => {
    doc.dispose();
  });

  it('removes the Nth paragraph, not the Nth body element, when a table precedes it', () => {
    // body: [Table, P0('KEEP-ME'), P1('REMOVE-ME')]
    doc.createTable(1, 1);
    doc.createParagraph('KEEP-ME');
    doc.createParagraph('REMOVE-ME');

    const removed = doc.removeParagraph(1);

    expect(removed).toBe(true);
    expect(doc.getParagraphCount()).toBe(1);
    expect(doc.getParagraphs()[0]?.getText()).toBe('KEEP-ME');
    expect(doc.getTableCount()).toBe(1);
  });

  it('round-trips with getParagraphIndex', () => {
    doc.createTable(1, 1);
    doc.createParagraph('first');
    const target = doc.createParagraph('target');
    doc.createParagraph('last');

    const removed = doc.removeParagraph(doc.getParagraphIndex(target));

    expect(removed).toBe(true);
    expect(doc.getParagraphs().map((p) => p.getText())).toEqual(['first', 'last']);
  });

  it('removes the same paragraph getParagraphAt resolves', () => {
    doc.createParagraph('p0');
    doc.createTable(1, 1);
    doc.createParagraph('p1');
    doc.createTable(1, 1);
    doc.createParagraph('p2');

    const expected = doc.getParagraphAt(2);
    expect(expected?.getText()).toBe('p2');

    expect(doc.removeParagraph(2)).toBe(true);
    expect(doc.getParagraphs().map((p) => p.getText())).toEqual(['p0', 'p1']);
  });

  it('returns false when the ordinal exceeds the paragraph count, even if bodyElements is longer', () => {
    doc.createTable(1, 1);
    doc.createParagraph('only');

    // bodyElements has 2 entries, but there is no paragraph at ordinal 1.
    expect(doc.removeParagraph(1)).toBe(false);
    expect(doc.getParagraphCount()).toBe(1);
    expect(doc.getTableCount()).toBe(1);
  });
});
