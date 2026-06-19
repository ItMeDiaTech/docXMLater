/**
 * Table.clone() must carry the table's own tracked-revision history.
 * w:tblPrChange (ECMA-376 §17.13.5.36) and w:tblGridChange (§17.13.5.35) live
 * in private fields outside formatting, so deep-cloning formatting alone
 * silently drops them — inconsistent with TableRow.clone() (trPrChange) and
 * TableCell.clone() (tcPrChange/cellRevision), and a clone of a parsed table
 * with pending table-property revisions would serialize without the markers.
 */
import { Table } from '../../src/elements/Table';
import { TableGridChange } from '../../src/elements/TableGridChange';

function findChild(xml: any, name: string): any {
  return (xml?.children ?? []).find((c: any) => c.name === name);
}

describe('C74: Table.clone() preserves tblPrChange and tblGridChange', () => {
  function buildTableWithRevisions(): Table {
    const table = new Table(2, 2);
    table.setTblPrChange({
      author: 'Reviewer',
      date: '2024-03-01T10:00:00Z',
      id: '7',
      previousProperties: { width: 5000 },
    });
    table.setTblGridChange(
      TableGridChange.create(8, [{ width: 2400 }, { width: 2400 }], 'Reviewer')
    );
    return table;
  }

  it('copies tblPrChange onto the clone', () => {
    const table = buildTableWithRevisions();
    const clone = table.clone();

    const change = clone.getTblPrChange();
    expect(change).toBeDefined();
    expect(change!.author).toBe('Reviewer');
    expect(change!.id).toBe('7');
    expect(change!.previousProperties.width).toBe(5000);
  });

  it('copies tblGridChange onto the clone', () => {
    const table = buildTableWithRevisions();
    const clone = table.clone();

    const gridChange = clone.getTblGridChange();
    expect(gridChange).toBeDefined();
    expect(gridChange!.getId()).toBe(8);
    expect(gridChange!.getAuthor()).toBe('Reviewer');
    expect(gridChange!.getPreviousGrid()).toEqual([{ width: 2400 }, { width: 2400 }]);
  });

  it('clones are independent copies, not shared references', () => {
    const table = buildTableWithRevisions();
    const clone = table.clone();

    expect(clone.getTblPrChange()).not.toBe(table.getTblPrChange());
    expect(clone.getTblGridChange()).not.toBe(table.getTblGridChange());

    // Mutating the clone's revision data must not leak into the original
    clone.getTblPrChange()!.previousProperties.width = 9999;
    expect(table.getTblPrChange()!.previousProperties.width).toBe(5000);

    clone.getTblGridChange()!.setId(99);
    expect(table.getTblGridChange()!.getId()).toBe(8);
  });

  it('serializes the revision markers from the clone', () => {
    const table = buildTableWithRevisions();
    const xml = table.clone().toXML();

    const tblPr = findChild(xml, 'w:tblPr');
    expect(findChild(tblPr, 'w:tblPrChange')).toBeDefined();

    const tblGrid = findChild(xml, 'w:tblGrid');
    expect(findChild(tblGrid, 'w:tblGridChange')).toBeDefined();
  });

  it('leaves revision fields unset when the original has none', () => {
    const clone = new Table(1, 1).clone();
    expect(clone.getTblPrChange()).toBeUndefined();
    expect(clone.getTblGridChange()).toBeUndefined();
  });
});
