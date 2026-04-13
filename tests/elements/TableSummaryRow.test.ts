/**
 * Tests for Table.addSummaryRow()
 */

import { Table } from '../../src/elements/Table';

describe('Table.addSummaryRow()', () => {
  describe('default behavior (sum)', () => {
    it('adds a totals row with summed numeric columns', () => {
      const table = Table.fromArray([
        ['Product', 'Price', 'Qty'],
        ['Widget', '10', '5'],
        ['Gadget', '25', '2'],
        ['Doohickey', '5', '8'],
      ]);

      table.addSummaryRow();

      expect(table.getRowCount()).toBe(5);
      const lastRow = table.getRow(4)!.toArray();
      expect(lastRow[0]).toBe('Total');
      expect(lastRow[1]).toBe('40'); // 10 + 25 + 5
      expect(lastRow[2]).toBe('15'); // 5 + 2 + 8
    });

    it('uses "Total" as default label', () => {
      const table = Table.fromArray([
        ['Name', 'Score'],
        ['Alice', '95'],
      ]);

      table.addSummaryRow();

      expect(table.getRow(2)!.toArray()[0]).toBe('Total');
    });

    it('returns empty string for non-numeric columns', () => {
      const table = Table.fromArray([
        ['Name', 'City', 'Score'],
        ['Alice', 'NYC', '95'],
        ['Bob', 'London', '87'],
      ]);

      table.addSummaryRow();

      const summary = table.getRow(3)!.toArray();
      expect(summary[0]).toBe('Total');
      expect(summary[1]).toBe(''); // Non-numeric
      expect(summary[2]).toBe('182'); // 95 + 87
    });

    it('handles table with only header (no data)', () => {
      const table = Table.fromArray([['Name', 'Value']]);

      table.addSummaryRow();

      const summary = table.getRow(1)!.toArray();
      expect(summary[0]).toBe('Total');
      expect(summary[1]).toBe(''); // No data to sum
    });

    it('handles decimal numbers', () => {
      const table = Table.fromArray([
        ['Item', 'Amount'],
        ['A', '10.5'],
        ['B', '20.3'],
      ]);

      table.addSummaryRow();

      const amount = parseFloat(table.getRow(3)!.toArray()[1]!);
      expect(amount).toBeCloseTo(30.8);
    });
  });

  describe('custom label', () => {
    it('uses custom label', () => {
      const table = Table.fromArray([
        ['Item', 'Count'],
        ['A', '10'],
      ]);

      table.addSummaryRow({ label: 'Grand Total' });

      expect(table.getRow(2)!.toArray()[0]).toBe('Grand Total');
    });
  });

  describe('custom startRow', () => {
    it('starts computation from a custom row', () => {
      const table = Table.fromArray([
        ['Category', 'Value'], // Row 0 (header)
        ['SubHeader', '---'], // Row 1 (skip this too)
        ['A', '10'], // Row 2 (start here)
        ['B', '20'], // Row 3
      ]);

      table.addSummaryRow({ startRow: 2 });

      expect(table.getRow(4)!.toArray()[1]).toBe('30'); // 10 + 20
    });
  });

  describe('custom compute function', () => {
    it('computes average instead of sum', () => {
      const table = Table.fromArray([
        ['Name', 'Score'],
        ['Alice', '90'],
        ['Bob', '80'],
        ['Charlie', '100'],
      ]);

      table.addSummaryRow({
        label: 'Average',
        compute: (values) => {
          const nums = values.map(Number).filter((n) => !isNaN(n));
          if (nums.length === 0) return '';
          return (nums.reduce((a, b) => a + b, 0) / nums.length).toFixed(1);
        },
      });

      expect(table.getRow(4)!.toArray()[0]).toBe('Average');
      expect(table.getRow(4)!.toArray()[1]).toBe('90.0');
    });

    it('computes count', () => {
      const table = Table.fromArray([
        ['Name', 'Status'],
        ['Alice', 'Active'],
        ['Bob', 'Inactive'],
        ['Charlie', 'Active'],
      ]);

      table.addSummaryRow({
        label: 'Count',
        compute: (values) => String(values.filter((v) => v.trim()).length),
      });

      expect(table.getRow(4)!.toArray()[1]).toBe('3');
    });

    it('receives column index', () => {
      const table = Table.fromArray([
        ['A', 'B', 'C'],
        ['1', '2', '3'],
      ]);

      const receivedIndices: number[] = [];
      table.addSummaryRow({
        compute: (values, colIndex) => {
          receivedIndices.push(colIndex);
          return '';
        },
      });

      // Column 0 uses label, so compute is called for columns 1, 2
      expect(receivedIndices).toEqual([1, 2]);
    });

    it('computes max per column', () => {
      const table = Table.fromArray([
        ['Name', 'Score'],
        ['Alice', '95'],
        ['Bob', '87'],
        ['Charlie', '92'],
      ]);

      table.addSummaryRow({
        label: 'Max',
        compute: (values) => {
          const nums = values.map(Number).filter((n) => !isNaN(n));
          return nums.length ? String(Math.max(...nums)) : '';
        },
      });

      expect(table.getRow(4)!.toArray()).toEqual(['Max', '95']);
    });
  });

  describe('returns created row', () => {
    it('returns the TableRow for further customization', () => {
      const table = Table.fromArray([
        ['A', 'B'],
        ['1', '2'],
      ]);

      const summaryRow = table.addSummaryRow();

      expect(summaryRow.toArray()[0]).toBe('Total');

      // Can customize the row further
      summaryRow.setHeader(false);
      for (const cell of summaryRow.getCells()) {
        cell.setBackgroundColor('E0E0E0');
      }

      expect(summaryRow.getCells()[0]!.getBackgroundColor()).toBe('E0E0E0');
    });
  });

  describe('single-column table', () => {
    it('adds label-only summary for single column', () => {
      const table = Table.fromArray([['Items'], ['A'], ['B']]);

      table.addSummaryRow();

      expect(table.getRow(3)!.toArray()).toEqual(['Total']);
    });
  });

  describe('integration with other table methods', () => {
    it('works with fromCSV + addSummaryRow', () => {
      const table = Table.fromCSV('Product,Revenue\nWidget,1000\nGadget,2500\nDoohickey,750');

      table.addSummaryRow();

      const summary = table.getRow(4)!.toArray();
      expect(summary[0]).toBe('Total');
      expect(summary[1]).toBe('4250');
    });

    it('works with mapColumn + addSummaryRow', () => {
      const table = Table.fromArray([
        ['Item', 'Price'],
        ['A', '10'],
        ['B', '20'],
      ]);

      // Format prices first
      table.mapColumn(1, (t, r) => (r === 0 ? t : `$${t}`));

      // Custom sum that strips $ prefix
      table.addSummaryRow({
        compute: (values) => {
          const nums = values.map((v) => parseFloat(v.replace('$', ''))).filter((n) => !isNaN(n));
          return nums.length ? `$${nums.reduce((a, b) => a + b, 0)}` : '';
        },
      });

      expect(table.getRow(3)!.toArray()[1]).toBe('$30');
    });

    it('toCSV includes summary row', () => {
      const table = Table.fromArray([
        ['Name', 'Score'],
        ['Alice', '90'],
        ['Bob', '80'],
      ]);

      table.addSummaryRow();

      const csv = table.toCSV();
      expect(csv).toContain('Total,170');
    });
  });
});
