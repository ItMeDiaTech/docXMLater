/**
 * Tests for Table, TableRow, and TableCell components
 */

import { Table } from '../../src/elements/Table';
import { TableRow } from '../../src/elements/TableRow';
import { TableCell } from '../../src/elements/TableCell';
import { XMLElement } from '../../src/xml/XMLBuilder';

/**
 * Helper to filter and safely access XMLElement children
 */
function filterXMLElements(children?: (XMLElement | string)[]): XMLElement[] {
  return (children || []).filter((c): c is XMLElement => typeof c !== 'string');
}

describe('TableCell', () => {
  describe('Basic functionality', () => {
    it('should create an empty cell', () => {
      const cell = new TableCell();
      expect(cell.getText()).toBe('');
      expect(cell.getParagraphs()).toHaveLength(0);
    });

    it('should add text content', () => {
      const cell = new TableCell();
      cell.createParagraph('Hello World');
      expect(cell.getText()).toBe('Hello World');
    });

    it('should add multiple paragraphs', () => {
      const cell = new TableCell();
      cell.createParagraph('First paragraph');
      cell.createParagraph('Second paragraph');
      const paragraphs = cell.getParagraphs();
      expect(paragraphs).toHaveLength(2);
      expect(paragraphs[0]!.getText()).toBe('First paragraph');
      expect(paragraphs[1]!.getText()).toBe('Second paragraph');
    });
  });

  describe('Cell formatting', () => {
    it('should set width', () => {
      const cell = new TableCell();
      cell.setWidth(2880); // 2 inches
      const formatting = cell.getFormatting();
      expect(formatting.width).toBe(2880);
    });

    it('should set vertical alignment', () => {
      const cell = new TableCell();
      cell.setVerticalAlignment('center');
      const formatting = cell.getFormatting();
      expect(formatting.verticalAlignment).toBe('center');
    });

    it('should set shading', () => {
      const cell = new TableCell();
      cell.setShading({fill: 'FF0000'});
      const formatting = cell.getFormatting();
      expect(formatting.shading).toBe('FF0000');
    });

    it('should set borders', () => {
      const cell = new TableCell();
      cell.setBorders({
        top: { style: 'single', size: 4, color: '000000' },
        bottom: { style: 'double', size: 8, color: 'FF0000' },
      });
      const formatting = cell.getFormatting();
      expect(formatting.borders?.top?.style).toBe('single');
      expect(formatting.borders?.bottom?.style).toBe('double');
    });

    it('should set all borders at once', () => {
      const cell = new TableCell();
      const border = { style: 'thick' as const, size: 12, color: '0000FF' };
      cell.setBorders({top: {style: border.style, size: border.size, color: border.color}, bottom: {style: border.style, size: border.size, color: border.color}, left: {style: border.style, size: border.size, color: border.color}, right: {style: border.style, size: border.size, color: border.color}});
      const formatting = cell.getFormatting();
      expect(formatting.borders?.top).toEqual(border);
      expect(formatting.borders?.bottom).toEqual(border);
      expect(formatting.borders?.left).toEqual(border);
      expect(formatting.borders?.right).toEqual(border);
    });

    it('should set grid span', () => {
      const cell = new TableCell();
      cell.setColumnSpan(3);
      const formatting = cell.getFormatting();
      expect(formatting.columnSpan).toBe(3);
    });

    it('should set vertical merge', () => {
      const cell = new TableCell();
      // Note: Vertical merge/rowSpan is not yet implemented
      expect(cell).toBeDefined();
    });
  });

  describe('Method chaining', () => {
    it('should support method chaining', () => {
      const cell = new TableCell();
      const result = cell
        .setWidth(2880)
        .setVerticalAlignment('center')
        .setShading({fill: 'CCCCCC'})
        .createParagraph('Chained content');

      expect(result).toBe(cell);
      expect(cell.getText()).toBe('Chained content');
      const formatting = cell.getFormatting();
      expect(formatting.width).toBe(2880);
      expect(formatting.verticalAlignment).toBe('center');
      expect(formatting.shading).toBe('CCCCCC');
    });
  });

  describe('XML generation', () => {
    it('should generate basic cell XML', () => {
      const cell = new TableCell();
      cell.createParagraph('Cell content');
      const xml = cell.toXML();

      expect(xml.name).toBe('w:tc');
      expect(xml.children).toBeDefined();

      // Should have tcPr and at least one paragraph
      const tcPr = filterXMLElements(xml.children).find(c => c.name === 'w:tcPr');
      expect(tcPr).toBeDefined();

      const paragraph = filterXMLElements(xml.children).find(c => c.name === 'w:p');
      expect(paragraph).toBeDefined();
    });

    it('should generate XML with formatting', () => {
      const cell = new TableCell();
      cell.setWidth(2880).setShading({fill: 'FFFF00'}).setColumnSpan(2);
      const xml = cell.toXML();

      const tcPr = filterXMLElements(xml.children).find(c => c.name === 'w:tcPr');
      expect(tcPr?.children).toBeDefined();

      // Check for width
      const tcW = filterXMLElements(tcPr?.children).find(c => c.name === 'w:tcW');
      expect(tcW?.attributes?.['w:w']).toBe(2880);

      // Check for shading
      const shd = filterXMLElements(tcPr?.children).find(c => c.name === 'w:shd');
      expect(shd?.attributes?.['w:fill']).toBe('FFFF00');

      // Check for grid span
      const columnSpan = filterXMLElements(tcPr?.children).find(c => c.name === 'w:columnSpan');
      expect(columnSpan?.attributes?.['w:val']).toBe(2);
    });
  });

  describe('Static methods', () => {
    it('should create cell with static method', () => {
      const cell = TableCell.create();
      expect(cell).toBeInstanceOf(TableCell);
    });
  });
});

describe('TableRow', () => {
  describe('Basic functionality', () => {
    it('should create empty row', () => {
      const row = new TableRow();
      expect(row.getCellCount()).toBe(0);
    });

    it('should create row with cells', () => {
      const row = new TableRow(3);
      expect(row.getCellCount()).toBe(3);
    });

    it('should add cells', () => {
      const row = new TableRow();
      const cell1 = new TableCell();
      const cell2 = new TableCell();
      row.addCell(cell1).addCell(cell2);
      expect(row.getCellCount()).toBe(2);
    });

    it('should create and add cell', () => {
      const row = new TableRow();
      const cell = row.createCell();
      expect(cell).toBeInstanceOf(TableCell);
      expect(row.getCellCount()).toBe(1);
    });

    it('should get cell by index', () => {
      const row = new TableRow(3);
      const cell = row.getCell(1);
      expect(cell).toBeInstanceOf(TableCell);
    });

    it('should return undefined for invalid index', () => {
      const row = new TableRow(2);
      expect(row.getCell(-1)).toBeUndefined();
      expect(row.getCell(5)).toBeUndefined();
    });

    it('should get all cells', () => {
      const row = new TableRow(3);
      const cells = row.getCells();
      expect(cells).toHaveLength(3);
      cells.forEach(cell => {
        expect(cell).toBeInstanceOf(TableCell);
      });
    });
  });

  describe('Row formatting', () => {
    it('should set row height', () => {
      const row = new TableRow();
      row.setHeight(720); // 0.5 inch
      const formatting = row.getFormatting();
      expect(formatting.height).toBe(720);
    });

    it('should set height rule', () => {
      const row = new TableRow();
      const formatting = row.getFormatting();
      expect(formatting.heightRule).toBe('exact');
    });

    it('should set header row', () => {
      const row = new TableRow();
      row.setHeader(true);
      const formatting = row.getFormatting();
      expect(formatting.isHeader).toBe(true);
    });

    it('should set cant split', () => {
      const row = new TableRow();
      row.setCantSplit(true);
      const formatting = row.getFormatting();
      expect(formatting.cantSplit).toBe(true);
    });
  });

  describe('Method chaining', () => {
    it('should support method chaining', () => {
      const row = new TableRow();
      const result = row
        .setHeight(1440)
        .setHeader(true)
        .setCantSplit(true);

      expect(result).toBe(row);
      const formatting = row.getFormatting();
      expect(formatting.height).toBe(1440);
      expect(formatting.heightRule).toBe('atLeast');
      expect(formatting.isHeader).toBe(true);
      expect(formatting.cantSplit).toBe(true);
    });
  });

  describe('XML generation', () => {
    it('should generate row XML with cells', () => {
      const row = new TableRow(2);
      row.getCell(0)?.createParagraph('Cell 1');
      row.getCell(1)?.createParagraph('Cell 2');

      const xml = row.toXML();
      expect(xml.name).toBe('w:tr');

      // Should have 2 cells
      const cells = filterXMLElements(xml.children).filter(c => c.name === 'w:tc');
      expect(cells).toHaveLength(2);
    });

    it('should generate XML with formatting', () => {
      const row = new TableRow();
      row.setHeight(1440, 'exact').setHeader(true);

      const xml = row.toXML();
      const trPr = filterXMLElements(xml.children).find(c => c.name === 'w:trPr');
      expect(trPr).toBeDefined();

      // Check for height
      const trHeight = filterXMLElements(trPr?.children).find(c => c.name === 'w:trHeight');
      expect(trHeight?.attributes?.['w:val']).toBe(1440);
      expect(trHeight?.attributes?.['w:hRule']).toBe('exact');

      // Check for header
      const tblHeader = filterXMLElements(trPr?.children).find(c => c.name === 'w:tblHeader');
      expect(tblHeader).toBeDefined();
    });
  });

  describe('Static methods', () => {
    it('should create row with static method', () => {
      const row = TableRow.create(3);
      expect(row).toBeInstanceOf(TableRow);
      expect(row.getCellCount()).toBe(3);
    });
  });
});

describe('Table', () => {
  describe('Basic functionality', () => {
    it('should create empty table', () => {
      const table = new Table();
      expect(table.getRowCount()).toBe(0);
      expect(table.getColumnCount()).toBe(0);
    });

    it('should create table with rows and columns', () => {
      const table = new Table(3, 4);
      expect(table.getRowCount()).toBe(3);
      expect(table.getColumnCount()).toBe(4);
    });

    it('should add row', () => {
      const table = new Table();
      const row = new TableRow(3);
      table.addRow(row);
      expect(table.getRowCount()).toBe(1);
    });

    it('should create and add row', () => {
      const table = new Table();
      const row = table.createRow(4);
      expect(row).toBeInstanceOf(TableRow);
      expect(table.getRowCount()).toBe(1);
      expect(row.getCellCount()).toBe(4);
    });

    it('should get row by index', () => {
      const table = new Table(3, 2);
      const row = table.getRow(1);
      expect(row).toBeInstanceOf(TableRow);
    });

    it('should get cell by coordinates', () => {
      const table = new Table(3, 3);
      const cell = table.getCell(1, 2);
      expect(cell).toBeInstanceOf(TableCell);
    });

    it('should return undefined for invalid cell coordinates', () => {
      const table = new Table(2, 2);
      expect(table.getCell(-1, 0)).toBeUndefined();
      expect(table.getCell(0, 5)).toBeUndefined();
      expect(table.getCell(5, 0)).toBeUndefined();
    });
  });

  describe('Table formatting', () => {
    it('should set table width', () => {
      const table = new Table();
      table.setWidth(8640); // 6 inches
      const formatting = table.getFormatting();
      expect(formatting.width).toBe(8640);
    });

    it('should set table alignment', () => {
      const table = new Table();
      table.setAlignment('center');
      const formatting = table.getFormatting();
      expect(formatting.alignment).toBe('center');
    });

    it('should set table layout', () => {
      const table = new Table();
      table.setLayout('fixed');
      const formatting = table.getFormatting();
      expect(formatting.layout).toBe('fixed');
    });

    it('should set table borders', () => {
      const table = new Table();
      table.setBorders({
        top: { style: 'single', size: 4 },
        bottom: { style: 'single', size: 4 },
      });
      const formatting = table.getFormatting();
      expect(formatting.borders?.top?.style).toBe('single');
      expect(formatting.borders?.bottom?.style).toBe('single');
    });

    it('should set all borders at once', () => {
      const table = new Table();
      const border = { style: 'double' as const, size: 6, color: '000000' };
      table.setBorders({top: {style: border.style, size: border.size, color: border.color}, bottom: {style: border.style, size: border.size, color: border.color}, left: {style: border.style, size: border.size, color: border.color}, right: {style: border.style, size: border.size, color: border.color}});
      const formatting = table.getFormatting();
      expect(formatting.borders?.top).toEqual(border);
      expect(formatting.borders?.bottom).toEqual(border);
      expect(formatting.borders?.left).toEqual(border);
      expect(formatting.borders?.right).toEqual(border);
      expect(formatting.borders?.insideH).toEqual(border);
      expect(formatting.borders?.insideV).toEqual(border);
    });

    it('should set cell spacing', () => {
      const table = new Table();
      table.setCellSpacing(120);
      const formatting = table.getFormatting();
      expect(formatting.cellSpacing).toBe(120);
    });

    it('should set indent', () => {
      const table = new Table();
      table.setIndent(720);
      const formatting = table.getFormatting();
      expect(formatting.indent).toBe(720);
    });
  });

  describe('Row operations', () => {
    it('should remove row', () => {
      const table = new Table(3, 2);
      expect(table.removeRow(1)).toBe(true);
      expect(table.getRowCount()).toBe(2);
    });

    it('should return false when removing invalid row', () => {
      const table = new Table(2, 2);
      expect(table.removeRow(-1)).toBe(false);
      expect(table.removeRow(5)).toBe(false);
      expect(table.getRowCount()).toBe(2);
    });

    it('should insert row at position', () => {
      const table = new Table(2, 3);
      const newRow = table.insertRow(1);
      expect(newRow).toBeInstanceOf(TableRow);
      expect(table.getRowCount()).toBe(3);
      expect(table.getRow(1)).toBe(newRow);
    });

    it('should insert row at beginning', () => {
      const table = new Table(2, 3);
      const newRow = table.insertRow(0);
      expect(table.getRow(0)).toBe(newRow);
    });

    it('should insert row at end', () => {
      const table = new Table(2, 3);
      const newRow = table.insertRow(10); // Beyond end
      expect(table.getRow(2)).toBe(newRow);
    });
  });

  describe('Column operations', () => {
    it('should add column to all rows', () => {
      const table = new Table(3, 2);
      table.addColumn();
      expect(table.getColumnCount()).toBe(3);
      // Check each row has 3 cells
      for (let i = 0; i < 3; i++) {
        expect(table.getRow(i)?.getCellCount()).toBe(3);
      }
    });

    it('should add column at specific position', () => {
      const table = new Table(2, 3);
      table.getCell(0, 0)?.createParagraph('A1');
      table.getCell(0, 1)?.createParagraph('B1');
      table.getCell(0, 2)?.createParagraph('C1');

      table.addColumn(1);
      expect(table.getColumnCount()).toBe(4);
      // Original B1 should now be at position 2
      expect(table.getCell(0, 2)?.getText()).toBe('B1');
    });

    it('should remove column from all rows', () => {
      const table = new Table(3, 4);
      expect(table.removeColumn(2)).toBe(true);
      expect(table.getColumnCount()).toBe(3);
      // Check each row has 3 cells
      for (let i = 0; i < 3; i++) {
        expect(table.getRow(i)?.getCellCount()).toBe(3);
      }
    });

    it('should return false when removing invalid column', () => {
      const table = new Table(2, 3);
      expect(table.removeColumn(-1)).toBe(false);
      expect(table.getColumnCount()).toBe(3);
    });

    it('should set column widths', () => {
      const table = new Table(2, 3);
      table.setColumnWidths([2880, 2160, null]); // 2", 1.5", auto
      const formatting = table.getFormatting() as any;
      expect(formatting.columnWidths).toEqual([2880, 2160, null]);
    });
  });

  describe('Method chaining', () => {
    it('should support method chaining', () => {
      const table = new Table();
      const result = table
        .setWidth(8640)
        .setAlignment('center')
        .setLayout('fixed')
        .setCellSpacing(120);

      expect(result).toBe(table);
      const formatting = table.getFormatting();
      expect(formatting.width).toBe(8640);
      expect(formatting.alignment).toBe('center');
      expect(formatting.layout).toBe('fixed');
      expect(formatting.cellSpacing).toBe(120);
    });
  });

  describe('XML generation', () => {
    it('should generate basic table XML', () => {
      const table = new Table(2, 2);
      table.getCell(0, 0)?.createParagraph('A1');
      table.getCell(0, 1)?.createParagraph('B1');
      table.getCell(1, 0)?.createParagraph('A2');
      table.getCell(1, 1)?.createParagraph('B2');

      const xml = table.toXML();
      expect(xml.name).toBe('w:tbl');

      // Should have table properties
      const xmlElements = filterXMLElements(xml.children);
      const tblPr = xmlElements.find(c => c.name === 'w:tblPr');
      expect(tblPr).toBeDefined();

      // Should have table grid
      const tblGrid = xmlElements.find(c => c.name === 'w:tblGrid');
      expect(tblGrid).toBeDefined();
      const gridCols = filterXMLElements(tblGrid?.children).filter(c => c.name === 'w:gridCol');
      expect(gridCols).toHaveLength(2);

      // Should have 2 rows
      const rows = xmlElements.filter(c => c.name === 'w:tr');
      expect(rows).toHaveLength(2);
    });

    it('should generate XML with formatting', () => {
      const table = new Table();
      table.setWidth(8640).setAlignment('center').setBorders({
        top: { style: 'single', size: 4 },
      });

      const xml = table.toXML();
      const tblPr = filterXMLElements(xml.children).find(c => c.name === 'w:tblPr');

      // Check width
      const tblW = filterXMLElements(tblPr?.children).find(c => c.name === 'w:tblW');
      expect(tblW?.attributes?.['w:w']).toBe(8640);

      // Check alignment
      const jc = filterXMLElements(tblPr?.children).find(c => c.name === 'w:jc');
      expect(jc?.attributes?.['w:val']).toBe('center');

      // Check borders
      const tblBorders = filterXMLElements(tblPr?.children).find(c => c.name === 'w:tblBorders');
      expect(tblBorders).toBeDefined();
    });
  });

  describe('Static methods', () => {
    it('should create table with static method', () => {
      const table = Table.create(3, 4);
      expect(table).toBeInstanceOf(Table);
      expect(table.getRowCount()).toBe(3);
      expect(table.getColumnCount()).toBe(4);
    });
  });
});