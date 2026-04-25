/**
 * InMemoryRevisionAcceptor — cellIns / cellDel / cellMerge routing.
 *
 * Per ECMA-376 Part 1 §17.13.5.4-6:
 *   - `w:cellIns` (§17.13.5.5) — tracked cell INSERTION.
 *   - `w:cellDel` (§17.13.5.6) — tracked cell DELETION.
 *   - `w:cellMerge` (§17.13.5.4) — tracked cell MERGE (property change).
 *
 * Bug guarded against: `acceptRevisionsInMemory` cleared all three markers
 * under `opts.acceptPropertyChanges`, regardless of the revision's semantic
 * type. A user who called `acceptAllRevisions({ acceptInsertions: true,
 * acceptDeletions: false })` would see `cellIns` markers persist (not
 * cleared because the gate was `acceptPropertyChanges`), and conversely
 * `{ acceptPropertyChanges: true }` alone would clear all three including
 * the insertions — neither matches the semantic contract of the options.
 *
 * Correct routing:
 *   - `tableCellInsert`  → `opts.acceptInsertions`,     `insertionsAccepted`
 *   - `tableCellDelete`  → `opts.acceptDeletions`,      `deletionsAccepted`
 *   - `tableCellMerge`   → `opts.acceptPropertyChanges`,`propertyChangesAccepted`
 */

import { Table } from '../../src/elements/Table';
import { TableRow } from '../../src/elements/TableRow';
import { TableCell } from '../../src/elements/TableCell';
import { Revision } from '../../src/elements/Revision';
import { Document } from '../../src/core/Document';
import { acceptRevisionsInMemory } from '../../src/processors/InMemoryRevisionAcceptor';

function buildDocWithCellRevisions(): Document {
  const doc = Document.create();
  const table = new Table(1, 3); // 1 row, 3 cols
  const row = table.getRows()[0]!;

  // Add text to each cell so the acceptor's empty-table sweep keeps it.
  row.getCells()[0]!.createParagraph('a');
  row.getCells()[1]!.createParagraph('b');
  row.getCells()[2]!.createParagraph('c');

  // Cell 0: tracked insertion
  row.getCells()[0]!.setCellRevision(Revision.createTableCellInsert('A', []));
  // Cell 1: tracked deletion
  row.getCells()[1]!.setCellRevision(Revision.createTableCellDelete('A', []));
  // Cell 2: tracked merge
  row.getCells()[2]!.setCellRevision(Revision.createTableCellMerge('A', []));

  doc.addBodyElement(table);
  return doc;
}

describe('acceptRevisionsInMemory — cellIns/cellDel/cellMerge type routing', () => {
  it('clears cellIns only when acceptInsertions is true', () => {
    const doc = buildDocWithCellRevisions();
    const result = acceptRevisionsInMemory(doc, {
      acceptInsertions: true,
      acceptDeletions: false,
      acceptMoves: false,
      acceptPropertyChanges: false,
    });
    const tableFound = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    const row = tableFound.getRows()[0]!;
    expect(row.getCells()[0]!.getCellRevision()).toBeUndefined();
    expect(row.getCells()[1]!.getCellRevision()).toBeDefined(); // cellDel untouched
    expect(row.getCells()[2]!.getCellRevision()).toBeDefined(); // cellMerge untouched
    expect(result.insertionsAccepted).toBeGreaterThanOrEqual(1);
    doc.dispose();
  });

  it('clears cellDel only when acceptDeletions is true', () => {
    const doc = buildDocWithCellRevisions();
    const result = acceptRevisionsInMemory(doc, {
      acceptInsertions: false,
      acceptDeletions: true,
      acceptMoves: false,
      acceptPropertyChanges: false,
    });
    const tableFound = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    const row = tableFound.getRows()[0]!;
    expect(row.getCells()[0]!.getCellRevision()).toBeDefined(); // cellIns untouched
    expect(row.getCells()[1]!.getCellRevision()).toBeUndefined();
    expect(row.getCells()[2]!.getCellRevision()).toBeDefined(); // cellMerge untouched
    expect(result.deletionsAccepted).toBeGreaterThanOrEqual(1);
    doc.dispose();
  });

  it('clears cellMerge only when acceptPropertyChanges is true', () => {
    const doc = buildDocWithCellRevisions();
    const result = acceptRevisionsInMemory(doc, {
      acceptInsertions: false,
      acceptDeletions: false,
      acceptMoves: false,
      acceptPropertyChanges: true,
    });
    const tableFound = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    const row = tableFound.getRows()[0]!;
    expect(row.getCells()[0]!.getCellRevision()).toBeDefined(); // cellIns untouched
    expect(row.getCells()[1]!.getCellRevision()).toBeDefined(); // cellDel untouched
    expect(row.getCells()[2]!.getCellRevision()).toBeUndefined();
    expect(result.propertyChangesAccepted).toBeGreaterThanOrEqual(1);
    doc.dispose();
  });

  it('clears all three when every flag is true', () => {
    const doc = buildDocWithCellRevisions();
    acceptRevisionsInMemory(doc, {
      acceptInsertions: true,
      acceptDeletions: true,
      acceptMoves: true,
      acceptPropertyChanges: true,
    });
    const tableFound = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    const row = tableFound.getRows()[0]!;
    expect(row.getCells()[0]!.getCellRevision()).toBeUndefined();
    expect(row.getCells()[1]!.getCellRevision()).toBeUndefined();
    expect(row.getCells()[2]!.getCellRevision()).toBeUndefined();
    doc.dispose();
  });

  it('increments the correct counter for each revision type', () => {
    const doc = buildDocWithCellRevisions();
    const result = acceptRevisionsInMemory(doc, {
      acceptInsertions: true,
      acceptDeletions: true,
      acceptMoves: true,
      acceptPropertyChanges: true,
    });
    // At minimum, each of the three markers is counted in its corresponding bucket.
    expect(result.insertionsAccepted).toBeGreaterThanOrEqual(1);
    expect(result.deletionsAccepted).toBeGreaterThanOrEqual(1);
    expect(result.propertyChangesAccepted).toBeGreaterThanOrEqual(1);
    doc.dispose();
  });
});
