/**
 * swapCells() must write through to each row's live cell array. The old
 * implementation swapped entries inside the defensive copies returned by
 * getCells(), so the table was never modified — a silent no-op.
 */
import { Document } from '../../src/core/Document';

describe('C17: swapCells writes through to the live cell arrays', () => {
  it('swaps cell content across rows', () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(2, 2);
      table.setCell(0, 0, 'A');
      table.setCell(1, 1, 'B');

      table.swapCells(0, 0, 1, 1);

      expect(table.getCell(0, 0)?.getText()).toBe('B');
      expect(table.getCell(1, 1)?.getText()).toBe('A');
    } finally {
      doc.dispose();
    }
  });

  it('updates each swapped cell parent-row reference', () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(2, 2);
      const cellA = table.getCell(0, 0)!;
      const cellB = table.getCell(1, 1)!;

      table.swapCells(0, 0, 1, 1);

      expect(cellA._getParentRow()).toBe(table.getRow(1));
      expect(cellB._getParentRow()).toBe(table.getRow(0));
    } finally {
      doc.dispose();
    }
  });

  it('swaps cells within the same row without orphaning either cell', () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(1, 2);
      table.setCell(0, 0, 'LEFT');
      table.setCell(0, 1, 'RIGHT');
      const left = table.getCell(0, 0)!;
      const right = table.getCell(0, 1)!;

      table.swapCells(0, 0, 0, 1);

      expect(table.getCell(0, 0)?.getText()).toBe('RIGHT');
      expect(table.getCell(0, 1)?.getText()).toBe('LEFT');
      expect(left._getParentRow()).toBe(table.getRow(0));
      expect(right._getParentRow()).toBe(table.getRow(0));
    } finally {
      doc.dispose();
    }
  });

  it('persists the swap through a save/reload round trip', async () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(2, 2);
      table.setCell(0, 0, 'A');
      table.setCell(1, 1, 'B');
      table.swapCells(0, 0, 1, 1);

      const buffer = await doc.toBuffer();
      const reloaded = await Document.loadFromBuffer(buffer);
      try {
        const reloadedTable = reloaded.getTables()[0]!;
        expect(reloadedTable.getCell(0, 0)?.getText()).toBe('B');
        expect(reloadedTable.getCell(1, 1)?.getText()).toBe('A');
      } finally {
        reloaded.dispose();
      }
    } finally {
      doc.dispose();
    }
  });
});
