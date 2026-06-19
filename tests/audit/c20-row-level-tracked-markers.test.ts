/**
 * Tracked whole-row operations must use the row-level trPr markers — w:del
 * (ECMA-376 §17.13.5.14) for deletions and w:ins (§17.13.5.19) for
 * insertions — not per-cell w:cellDel/w:cellIns, which track cell-structure
 * changes and are never resolved into a row removal by the revision
 * acceptor. Column deletions legitimately keep the cell-level marker.
 */
import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function getDocumentXml(doc: Document): Promise<string> {
  const buffer = await doc.toBuffer();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  return zip.getFileAsString('word/document.xml')!;
}

// Matches a marker element that lives INSIDE a w:trPr block
function trPrMarker(name: 'ins' | 'del'): RegExp {
  return new RegExp(`<w:trPr>(?:(?!</w:trPr>)[^])*<w:${name} [^>]*w:author="TestAuthor"`);
}

describe('Row-level tracked markers for row operations', () => {
  let doc: Document;

  afterEach(() => {
    doc.dispose();
  });

  it('removeRow marks the row with trPr w:del and accept removes the row', async () => {
    doc = Document.create();
    const table = new Table(3, 2);
    doc.addTable(table);
    table.getRows()[1]!.getCells()[0]!.createParagraph('Doomed');
    doc.enableTrackChanges({ author: 'TestAuthor' });

    expect(table.removeRow(1)).toBe(true);
    expect(table.getRows().length).toBe(3);

    const row = table.getRows()[1]!;
    const del = row.getRowDeletion();
    expect(del).toBeDefined();
    expect(del!.author).toBe('TestAuthor');
    for (const cell of row.getCells()) {
      expect(cell.getCellRevision()).toBeUndefined();
    }

    const xml = await getDocumentXml(doc);
    expect(xml).toMatch(trPrMarker('del'));

    await doc.acceptAllRevisions();
    expect(table.getRows().length).toBe(2);
    expect(table.getRows().some((r) => r.getRowDeletion() !== undefined)).toBe(false);
  });

  it('removeRows marks every removed row and accept splices them all', async () => {
    doc = Document.create();
    const table = new Table(4, 1);
    doc.addTable(table);
    doc.enableTrackChanges({ author: 'TestAuthor' });

    expect(table.removeRows(1, 2)).toBe(true);
    expect(table.getRows().length).toBe(4);
    expect(table.getRows()[1]!.getRowDeletion()).toBeDefined();
    expect(table.getRows()[2]!.getRowDeletion()).toBeDefined();
    expect(table.getRows()[0]!.getRowDeletion()).toBeUndefined();
    expect(table.getRows()[3]!.getRowDeletion()).toBeUndefined();

    await doc.acceptAllRevisions();
    expect(table.getRows().length).toBe(2);
  });

  it('insertRow marks the row with trPr w:ins; accept keeps the row and clears the marker', async () => {
    doc = Document.create();
    const table = new Table(2, 2);
    doc.addTable(table);
    doc.enableTrackChanges({ author: 'TestAuthor' });

    const row = table.insertRow(1);
    expect(table.getRows().length).toBe(3);
    const ins = row.getRowInsertion();
    expect(ins).toBeDefined();
    expect(ins!.author).toBe('TestAuthor');
    for (const cell of row.getCells()) {
      expect(cell.getCellRevision()).toBeUndefined();
    }

    const xml = await getDocumentXml(doc);
    expect(xml).toMatch(trPrMarker('ins'));

    await doc.acceptAllRevisions();
    expect(table.getRows().length).toBe(3);
    expect(row.getRowInsertion()).toBeUndefined();
  });

  it('insertRows marks every inserted row with trPr w:ins', () => {
    doc = Document.create();
    const table = new Table(2, 1);
    doc.addTable(table);
    doc.enableTrackChanges({ author: 'TestAuthor' });

    const rows = table.insertRows(1, 2);
    expect(rows.length).toBe(2);
    for (const row of rows) {
      expect(row.getRowInsertion()).toBeDefined();
    }
  });

  it('removeColumn keeps cell-level cellDel markers and wraps the text once', async () => {
    doc = Document.create();
    const table = new Table(2, 2);
    doc.addTable(table);
    table.getRows()[0]!.getCells()[1]!.createParagraph('ColText');
    doc.enableTrackChanges({ author: 'TestAuthor' });

    expect(table.removeColumn(1)).toBe(true);
    for (const row of table.getRows()) {
      expect(row.getRowDeletion()).toBeUndefined();
      const cell = row.getCells()[1]!;
      expect(cell.getCellRevision()).toBeDefined();
      expect(cell.getCellRevision()!.getType()).toBe('tableCellDelete');
    }

    const xml = await getDocumentXml(doc);
    expect((xml.match(/ColText/g) || []).length).toBe(1);
    expect(xml).toMatch(/<w:delText[^>]*>ColText/);
  });

  it('row-level markers survive a save/load round-trip', async () => {
    doc = Document.create();
    const table = new Table(3, 1);
    doc.addTable(table);
    doc.enableTrackChanges({ author: 'TestAuthor' });
    table.removeRow(1);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    try {
      const loadedRow = loaded.getTables()[0]!.getRows()[1]!;
      const del = loadedRow.getRowDeletion();
      expect(del).toBeDefined();
      expect(del!.author).toBe('TestAuthor');
    } finally {
      loaded.dispose();
    }
  });
});
