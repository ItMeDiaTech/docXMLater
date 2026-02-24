/**
 * Tests for diagonal cell borders (tl2br, tr2bl)
 * and cell headers attribute
 */

import { Table } from '../../src/elements/Table';
import { TableCell } from '../../src/elements/TableCell';
import { Document } from '../../src/core/Document';
import { XMLElement } from '../../src/xml/XMLBuilder';

function filterXMLElements(children?: (XMLElement | string)[]): XMLElement[] {
  return (children || []).filter((c): c is XMLElement => typeof c !== 'string');
}

describe('Table Cell Diagonal Borders', () => {
  describe('Top-Left to Bottom-Right (tl2br)', () => {
    test('should set tl2br border', () => {
      const cell = new TableCell();
      cell.setBorders({
        tl2br: { style: 'single', size: 4, color: '000000' },
      });
      expect(cell.getBorders()?.tl2br?.style).toBe('single');
    });

    test('should generate w:tl2br in XML', () => {
      const cell = new TableCell();
      cell.setBorders({
        tl2br: { style: 'single', size: 4, color: 'FF0000' },
      });
      cell.createParagraph('Diagonal');

      const xml = cell.toXML();
      const tcPr = filterXMLElements(xml.children).find((c) => c.name === 'w:tcPr');
      const tcBorders = filterXMLElements(tcPr?.children).find((c) => c.name === 'w:tcBorders');
      const tl2br = filterXMLElements(tcBorders?.children).find((c) => c.name === 'w:tl2br');

      expect(tl2br).toBeDefined();
      expect(tl2br?.attributes?.['w:val']).toBe('single');
      expect(tl2br?.attributes?.['w:color']).toBe('FF0000');
    });

    test('should round-trip tl2br border', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table
        .getRow(0)!
        .getCell(0)!
        .setBorders({
          tl2br: { style: 'single', size: 8, color: '0000FF' },
        });
      table.getRow(0)!.getCell(0)!.createParagraph('Diagonal TL-BR');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const borders = loaded.getTables()[0]!.getRow(0)!.getCell(0)!.getBorders();

      expect(borders?.tl2br).toBeDefined();
      expect(borders?.tl2br?.style).toBe('single');
      expect(borders?.tl2br?.size).toBe(8);
      expect(borders?.tl2br?.color).toBe('0000FF');

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Top-Right to Bottom-Left (tr2bl)', () => {
    test('should set tr2bl border', () => {
      const cell = new TableCell();
      cell.setBorders({
        tr2bl: { style: 'double', size: 6, color: 'FF0000' },
      });
      expect(cell.getBorders()?.tr2bl?.style).toBe('double');
    });

    test('should round-trip tr2bl border', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table
        .getRow(0)!
        .getCell(0)!
        .setBorders({
          tr2bl: { style: 'double', size: 6, color: 'FF0000' },
        });
      table.getRow(0)!.getCell(0)!.createParagraph('Diagonal TR-BL');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const borders = loaded.getTables()[0]!.getRow(0)!.getCell(0)!.getBorders();

      expect(borders?.tr2bl).toBeDefined();
      expect(borders?.tr2bl?.style).toBe('double');

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Both Diagonal Borders', () => {
    test('should round-trip both diagonal borders', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table
        .getRow(0)!
        .getCell(0)!
        .setBorders({
          top: { style: 'single', size: 4, color: '000000' },
          bottom: { style: 'single', size: 4, color: '000000' },
          left: { style: 'single', size: 4, color: '000000' },
          right: { style: 'single', size: 4, color: '000000' },
          tl2br: { style: 'single', size: 4, color: 'FF0000' },
          tr2bl: { style: 'single', size: 4, color: '0000FF' },
        });
      table.getRow(0)!.getCell(0)!.createParagraph('X borders');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const borders = loaded.getTables()[0]!.getRow(0)!.getCell(0)!.getBorders();

      expect(borders?.tl2br?.color).toBe('FF0000');
      expect(borders?.tr2bl?.color).toBe('0000FF');
      expect(borders?.top?.style).toBe('single');

      doc.dispose();
      loaded.dispose();
    });
  });
});

describe('Table Cell Headers Attribute', () => {
  test('should set and get headers', () => {
    const cell = new TableCell();
    cell.setHeaders('header1 header2');
    expect(cell.getHeaders()).toBe('header1 header2');
  });

  test('should not generate w:headers in XML (not in Transitional schema)', () => {
    // w:headers is in ECMA-376 Strict but NOT in Transitional schema.
    // We store it in memory but do not serialize it to avoid OOXML validation errors.
    const cell = new TableCell();
    cell.setHeaders('hdr_name hdr_date');
    cell.createParagraph('Data');

    const xml = cell.toXML();
    const tcPr = filterXMLElements(xml.children).find((c) => c.name === 'w:tcPr');
    const headers = filterXMLElements(tcPr?.children).find((c) => c.name === 'w:headers');

    expect(headers).toBeUndefined();
    // But the in-memory value is still accessible
    expect(cell.getHeaders()).toBe('hdr_name hdr_date');
  });

  test('should preserve headers in memory across set/get', () => {
    const cell = new TableCell();
    cell.setHeaders('col1_header');
    cell.createParagraph('Data Cell');
    expect(cell.getHeaders()).toBe('col1_header');

    // Update headers
    cell.setHeaders('updated_header');
    expect(cell.getHeaders()).toBe('updated_header');
  });
});

describe('Table Row DivId', () => {
  test('should set and get divId on row', () => {
    const table = new Table(2, 2);
    table.getRow(0)!.setDivId(42);
    expect(table.getRow(0)!.getDivId()).toBe(42);
  });

  test('should round-trip row divId', async () => {
    const doc = Document.create();
    const table = new Table(2, 2);
    table.getRow(0)!.setDivId(12345);
    table.getRow(0)!.getCell(0)!.createParagraph('With DivId');
    doc.addTable(table);

    const buffer = await doc.toBuffer();
    const loaded = await Document.loadFromBuffer(buffer);
    expect(loaded.getTables()[0]!.getRow(0)!.getDivId()).toBe(12345);

    doc.dispose();
    loaded.dispose();
  });
});
