/**
 * Document.replaceFormattedText must visit each table-cell run exactly once.
 *
 * getAllParagraphs() walks recursively and already includes every table-cell
 * paragraph, so the previous extra getTables() loop re-applied the regex to
 * the same Run instances. Whenever the replacement still matched the search
 * pattern ('cat' -> 'cats'), table text was replaced twice ('catss') and the
 * returned count was inflated.
 */

import { Document } from '../../src/core/Document';

describe('replaceFormattedText visits table-cell runs once', () => {
  let doc: Document;

  beforeEach(() => {
    doc = Document.create();
  });

  afterEach(() => {
    doc.dispose();
  });

  it('replaces table-cell text once when the replacement rematches the pattern', () => {
    const table = doc.createTable(1, 1);
    table.getCell(0, 0)!.createParagraph('cat');

    const count = doc.replaceFormattedText('cat', 'cats');

    expect(count).toBe(1);
    expect(table.getCell(0, 0)?.getText()).toBe('cats');
  });

  it('counts body and table occurrences exactly once each', () => {
    doc.createParagraph('cat');
    const table = doc.createTable(1, 1);
    table.getCell(0, 0)!.createParagraph('cat');

    const count = doc.replaceFormattedText('cat', 'cats');

    expect(count).toBe(2);
    expect(doc.getParagraphs()[1]?.getText()).toBe('cats');
    expect(table.getCell(0, 0)?.getText()).toBe('cats');
  });

  it('persists the single replacement on save/reload', async () => {
    const table = doc.createTable(1, 1);
    table.getCell(0, 0)!.createParagraph('cat');

    doc.replaceFormattedText('cat', 'cats');

    const buffer = await doc.toBuffer();
    const reloaded = await Document.loadFromBuffer(buffer);
    try {
      expect(reloaded.getTables()[0]?.getCell(0, 0)?.getText()).toBe('cats');
    } finally {
      reloaded.dispose();
    }
  });
});
