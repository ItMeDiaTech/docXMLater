/**
 * Table.addColumn() must insert a numeric width into w:tblGrid even when the
 * grid has no usable neighbor (e.g. an empty <w:tblGrid/>).
 *
 * Previously, when both the left and right neighbor were undefined, the
 * computed width was `undefined` and was spliced into formatting.tableGrid as
 * `undefined as number`, producing a corrupt grid entry that serializes to an
 * invalid <w:gridCol w:w="undefined"/>. The fix falls back to 1 inch (1440
 * twips) so a numeric width is always inserted.
 */

import { Table } from '../../src/elements/Table';

describe('addColumn fills an empty tableGrid with a numeric width', () => {
  it('inserts 1440 twips when the grid has no usable neighbor', () => {
    const table = new Table(1, 1);
    table.setTableGrid([]); // truthy but empty grid

    table.addColumn();

    const grid = table.getTableGrid();
    expect(grid).toBeDefined();
    expect(grid!.length).toBe(1);
    expect(typeof grid![0]).toBe('number');
    expect(Number.isFinite(grid![0])).toBe(true);
    expect(grid![0]).toBe(1440);
  });

  it('mirrors an existing neighbor width when one is present', () => {
    const table = new Table(1, 1);
    table.setTableGrid([2880]);

    table.addColumn(); // append at end → mirrors the single existing column

    const grid = table.getTableGrid();
    expect(grid!.length).toBe(2);
    expect(grid!.every((w) => typeof w === 'number' && Number.isFinite(w))).toBe(true);
    expect(grid![1]).toBe(2880);
  });
});
