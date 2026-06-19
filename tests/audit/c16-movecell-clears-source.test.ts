/**
 * moveCell() must actually clear the source cell. The old implementation
 * assigned a new TableCell into the defensive copy returned by getCells(),
 * so the source kept its content (serialized twice) and the same Paragraph
 * instances lived in both cells at once.
 */
import { Document } from '../../src/core/Document';

describe('C16: moveCell clears the source cell', () => {
  it('moves content to the target and leaves the source cell empty', () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(2, 2);
      table.setCell(0, 0, 'MOVE');
      const sourceCell = table.getCell(0, 0)!;

      table.moveCell(0, 0, 1, 1);

      expect(table.getCell(1, 1)?.getText()).toContain('MOVE');
      expect(table.getCell(0, 0)?.getText()).toBe('');
      // The moved paragraphs must leave the source cell entirely so no
      // Paragraph instance is shared between two cells
      expect(sourceCell.getParagraphs()).toHaveLength(0);
    } finally {
      doc.dispose();
    }
  });

  it('replaces the source cell in the row (not in a throwaway copy)', () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(2, 2);
      table.setCell(0, 0, 'MOVE');
      const originalSource = table.getCell(0, 0);

      table.moveCell(0, 0, 1, 1);

      expect(table.getCell(0, 0)).not.toBe(originalSource);
    } finally {
      doc.dispose();
    }
  });

  it('serializes the moved content exactly once across a round trip', async () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(2, 2);
      table.setCell(0, 0, 'MOVE');
      table.moveCell(0, 0, 1, 1);

      const buffer = await doc.toBuffer();
      const reloaded = await Document.loadFromBuffer(buffer);
      try {
        const reloadedTable = reloaded.getTables()[0]!;
        expect(reloadedTable.getCell(0, 0)?.getText()).toBe('');
        expect(reloadedTable.getCell(1, 1)?.getText()).toContain('MOVE');
      } finally {
        reloaded.dispose();
      }
    } finally {
      doc.dispose();
    }
  });

  it('is a no-op when source and target are the same cell', () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(2, 2);
      table.setCell(0, 0, 'SAME');

      table.moveCell(0, 0, 0, 0);

      expect(table.getCell(0, 0)?.getText()).toBe('SAME');
      expect(table.getCell(0, 0)?.getParagraphs()).toHaveLength(1);
    } finally {
      doc.dispose();
    }
  });
});
