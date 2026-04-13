/**
 * Tests for Table.transpose()
 */

import { Table } from '../../src/elements/Table';
import { TableRow } from '../../src/elements/TableRow';
import { TableCell } from '../../src/elements/TableCell';

describe('Table.transpose()', () => {
  describe('basic transposition', () => {
    it('transposes a 2x3 table to 3x2', () => {
      const table = Table.fromArray([
        ['A1', 'A2', 'A3'],
        ['B1', 'B2', 'B3'],
      ]);

      const transposed = table.transpose();

      expect(transposed.getRowCount()).toBe(3);
      expect(transposed.getColumnCount()).toBe(2);
      expect(transposed.toArray()).toEqual([
        ['A1', 'B1'],
        ['A2', 'B2'],
        ['A3', 'B3'],
      ]);
    });

    it('transposes a 3x2 table to 2x3', () => {
      const table = Table.fromArray([
        ['Name', 'Age'],
        ['Alice', '30'],
        ['Bob', '25'],
      ]);

      const transposed = table.transpose();

      expect(transposed.toArray()).toEqual([
        ['Name', 'Alice', 'Bob'],
        ['Age', '30', '25'],
      ]);
    });

    it('transposes a 1x1 table (identity)', () => {
      const table = Table.fromArray([['Solo']]);
      const transposed = table.transpose();

      expect(transposed.toArray()).toEqual([['Solo']]);
    });

    it('transposes a single row to a single column', () => {
      const table = Table.fromArray([['A', 'B', 'C']]);
      const transposed = table.transpose();

      expect(transposed.toArray()).toEqual([['A'], ['B'], ['C']]);
    });

    it('transposes a single column to a single row', () => {
      const table = Table.fromArray([['X'], ['Y'], ['Z']]);
      const transposed = table.transpose();

      expect(transposed.toArray()).toEqual([['X', 'Y', 'Z']]);
    });
  });

  describe('data integrity', () => {
    it('does not modify the original table', () => {
      const original = [
        ['A', 'B'],
        ['C', 'D'],
      ];
      const table = Table.fromArray(original);
      table.transpose();

      expect(table.toArray()).toEqual(original);
    });

    it('double transpose restores original data', () => {
      const original = [
        ['A1', 'A2', 'A3'],
        ['B1', 'B2', 'B3'],
      ];
      const table = Table.fromArray(original);

      const result = table.transpose().transpose();

      expect(result.toArray()).toEqual(original);
    });

    it('preserves cell content with multiple paragraphs', () => {
      const table = new Table(1, 2);
      const cell = table.getCell(0, 0)!;
      cell.createParagraph('Line 1');
      cell.createParagraph('Line 2');
      table.getCell(0, 1)!.createParagraph('Simple');

      const transposed = table.transpose();

      // Cell with multi-line content should be preserved
      expect(transposed.getCell(0, 0)!.getText()).toBe('Line 1\nLine 2');
      expect(transposed.getCell(1, 0)!.getText()).toBe('Simple');
    });
  });

  describe('cell formatting preservation', () => {
    it('clones cell formatting during transpose', () => {
      const table = new Table(1, 2);
      table.getCell(0, 0)!.setShading({ fill: 'FF0000', pattern: 'clear' });
      table.getCell(0, 0)!.createParagraph('Red');
      table.getCell(0, 1)!.createParagraph('Plain');

      const transposed = table.transpose();

      expect(transposed.getCell(0, 0)!.getFormatting().shading?.fill).toBe('FF0000');
      expect(transposed.getCell(0, 0)!.getText()).toBe('Red');
    });

    it('formatting is independent after transpose', () => {
      const table = new Table(1, 2);
      table.getCell(0, 0)!.setWidth(2400);
      table.getCell(0, 0)!.createParagraph('A');
      table.getCell(0, 1)!.createParagraph('B');

      const transposed = table.transpose();
      transposed.getCell(0, 0)!.setWidth(4800);

      expect(table.getCell(0, 0)!.getFormatting().width).toBe(2400);
    });
  });

  describe('edge cases', () => {
    it('handles empty table', () => {
      const table = new Table(0, 0);
      const transposed = table.transpose();

      expect(transposed.getRowCount()).toBe(0);
    });

    it('handles square table', () => {
      const table = Table.fromArray([
        ['1', '2', '3'],
        ['4', '5', '6'],
        ['7', '8', '9'],
      ]);

      const transposed = table.transpose();

      expect(transposed.toArray()).toEqual([
        ['1', '4', '7'],
        ['2', '5', '8'],
        ['3', '6', '9'],
      ]);
    });

    it('generates valid XML', () => {
      const table = Table.fromArray([
        ['A', 'B'],
        ['C', 'D'],
      ]);

      const transposed = table.transpose();
      const xml = transposed.toXML();

      expect(xml.name).toBe('w:tbl');
      const rows = xml.children!.filter((c) => typeof c !== 'string' && c.name === 'w:tr');
      expect(rows).toHaveLength(2);
    });
  });

  describe('practical use cases', () => {
    it('converts row-oriented data to column-oriented', () => {
      // Common pattern: metrics in rows → metrics in columns
      const metrics = Table.fromArray([
        ['Metric', 'Q1', 'Q2', 'Q3', 'Q4'],
        ['Revenue', '100', '120', '130', '150'],
        ['Costs', '80', '85', '90', '95'],
      ]);

      const columnar = metrics.transpose();

      expect(columnar.getRowCount()).toBe(5);
      expect(columnar.getColumnCount()).toBe(3);
      expect(columnar.getCell(0, 0)!.getText()).toBe('Metric');
      expect(columnar.getCell(1, 0)!.getText()).toBe('Q1');
      expect(columnar.getCell(1, 1)!.getText()).toBe('100');
      expect(columnar.getCell(1, 2)!.getText()).toBe('80');
    });
  });
});
