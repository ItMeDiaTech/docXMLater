/**
 * Word repairs files where a nested w:tbl is the final child of w:tc — every
 * cell must end with w:p. The old trailing-paragraph guard in
 * TableCell.toXML() only fired when the cell had zero paragraphs, and
 * removeParagraph() could strip the required paragraph that followed a nested
 * table (e.g. via Table.setCell on a nested-table cell).
 */
import { Document } from '../../src/core/Document';
import { TableCell } from '../../src/elements/TableCell';
import { ZipHandler } from '../../src/zip/ZipHandler';

const NESTED_TABLE_XML =
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';

function lastChildName(cell: TableCell): string | undefined {
  const children = cell.toXML().children ?? [];
  const last = children[children.length - 1];
  if (last && typeof last === 'object' && 'name' in last) {
    return last.name;
  }
  return undefined;
}

describe('C21: cell never serializes with a nested table as its last child', () => {
  it('toXML appends a trailing paragraph when raw nested content ends the cell', () => {
    const cell = new TableCell();
    cell.createParagraph('Before');
    cell.addRawNestedContent(1, NESTED_TABLE_XML, 'table');

    expect(lastChildName(cell)).toBe('w:p');
  });

  it('toXML does not add an extra paragraph when one already follows the nested table', () => {
    const cell = new TableCell();
    cell.createParagraph('Before');
    cell.addRawNestedContent(1, NESTED_TABLE_XML, 'table');
    cell.createParagraph('After');

    const children = cell.toXML().children ?? [];
    const paragraphCount = children.filter(
      (child) => typeof child === 'object' && 'name' in child && child.name === 'w:p'
    ).length;
    expect(paragraphCount).toBe(2);
    expect(lastChildName(cell)).toBe('w:p');
  });

  it('removeParagraph refuses to strip the paragraph that closes a nested table', () => {
    const cell = new TableCell();
    cell.createParagraph('Before');
    cell.createParagraph('After');
    // Shape: Before, nested table, After
    cell.addRawNestedContent(1, NESTED_TABLE_XML, 'table');

    expect(cell.removeParagraph(1)).toBe(false);
    expect(cell.getParagraphs()).toHaveLength(2);
  });

  it('setCell on a nested-table cell still saves a w:p as the last child of w:tc', async () => {
    const doc = Document.create();
    try {
      const table = doc.createTable(1, 1);
      const cell = table.getCell(0, 0)!;
      cell.createParagraph('Trailing');
      // Shape: empty paragraph, nested table, Trailing
      cell.addRawNestedContent(1, NESTED_TABLE_XML, 'table');

      table.setCell(0, 0, 'Updated');

      const buffer = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const documentXml = zip.getFileAsString('word/document.xml');
      // The outer cell must not close immediately after the nested table
      expect(documentXml).not.toContain('</w:tbl></w:tc>');
      expect(documentXml).toContain('Updated');
    } finally {
      doc.dispose();
    }
  });
});
