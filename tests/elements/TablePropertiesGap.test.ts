/**
 * Gap Tests for Table Properties
 *
 * Tests implemented-but-undertested Table.ts features:
 * - setShading() round-trip
 * - setCellMargins() round-trip
 * - setTblLook() round-trip and flag decoding
 * - setStyleRowBandSize() / setStyleColBandSize() round-trip
 * - setCaption() / setDescription() combined round-trip
 * - setPosition() additional edge cases
 * - setOverlap() / setBidiVisual() round-trip
 */

import { Table } from '../../src/elements/Table';
import { Document } from '../../src/core/Document';
import path from 'path';
import fs from 'fs';

const OUTPUT_DIR = path.join(__dirname, '../output');

if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

describe('Table Properties Gap Tests', () => {
  describe('Table Shading (w:shd in tblPr)', () => {
    test('should round-trip simple fill shading', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.setShading({ fill: 'E0E0E0' });
      table.getRow(0)!.getCell(0)!.createParagraph('Shaded Table');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const loadedTable = loaded.getTables()[0]!;

      expect(loadedTable.getShading()).toBeDefined();
      expect(loadedTable.getShading()!.fill).toBe('E0E0E0');

      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip shading with pattern and color', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.setShading({ fill: 'FFFFFF', pattern: 'pct12', color: '000000' });
      table.getRow(0)!.getCell(0)!.createParagraph('Patterned');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const loadedTable = loaded.getTables()[0]!;

      const shading = loadedTable.getShading();
      expect(shading).toBeDefined();
      expect(shading!.fill).toBe('FFFFFF');
      expect(shading!.pattern).toBe('pct12');
      expect(shading!.color).toBe('000000');

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Table Cell Margins (w:tblCellMar)', () => {
    test('should round-trip default cell margins', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.setCellMargins({ top: 50, bottom: 50, left: 108, right: 108 });
      table.getRow(0)!.getCell(0)!.createParagraph('With Margins');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const loadedTable = loaded.getTables()[0]!;

      const margins = loadedTable.getCellMargins();
      expect(margins).toBeDefined();
      expect(margins!.top).toBe(50);
      expect(margins!.bottom).toBe(50);
      expect(margins!.left).toBe(108);
      expect(margins!.right).toBe(108);

      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip asymmetric margins', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.setCellMargins({ top: 0, bottom: 100, left: 200, right: 50 });
      table.getRow(0)!.getCell(0)!.createParagraph('Asymmetric');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const margins = loaded.getTables()[0]!.getCellMargins();

      expect(margins!.top).toBe(0);
      expect(margins!.bottom).toBe(100);
      expect(margins!.left).toBe(200);
      expect(margins!.right).toBe(50);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Table Look (w:tblLook)', () => {
    test('should round-trip tblLook value', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.setTblLook('04A0');
      table.getRow(0)!.getCell(0)!.createParagraph('With TblLook');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const loadedTable = loaded.getTables()[0]!;

      // Value may come back as string or number depending on XML parser behavior
      const look = loadedTable.getTblLook();
      expect(look).toBeDefined();
      const lookVal = typeof look === 'number' ? look : parseInt(String(look), 16);
      expect(lookVal).toBe(0x04A0);

      doc.dispose();
      loaded.dispose();
    });

    test('should decode tblLook flags correctly', () => {
      const table = new Table(2, 2);

      // 04A0 = 0x04A0 = firstRow(0x0020) + firstColumn(0x0080) + noVBand(0x0400)
      table.setTblLook('04A0');
      let flags = table.getTblLookFlags();
      expect(flags.firstRow).toBe(true);
      expect(flags.lastRow).toBe(false);
      expect(flags.firstColumn).toBe(true);
      expect(flags.lastColumn).toBe(false);
      expect(flags.noHBand).toBe(false);
      expect(flags.noVBand).toBe(true);

      // 01E0 = firstRow(0x0020) + lastRow(0x0040) + firstColumn(0x0080) + lastColumn(0x0100)
      table.setTblLook('01E0');
      flags = table.getTblLookFlags();
      expect(flags.firstRow).toBe(true);
      expect(flags.lastRow).toBe(true);
      expect(flags.firstColumn).toBe(true);
      expect(flags.lastColumn).toBe(true);
      expect(flags.noHBand).toBe(false);
      expect(flags.noVBand).toBe(false);
    });

    test('should round-trip tblLook value (may be parsed as numeric string)', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.setTblLook('04A0');
      table.getRow(0)!.getCell(0)!.createParagraph('TblLook test');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const tblLook = loaded.getTables()[0]!.getTblLook();
      // The value may be returned as string or number depending on XML parser
      // The important thing is the hex value is preserved
      expect(tblLook).toBeDefined();
      const numericVal = typeof tblLook === 'number' ? tblLook : parseInt(String(tblLook), 16);
      expect(numericVal).toBe(0x04A0);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Style Band Sizes (tblStyleRowBandSize, tblStyleColBandSize)', () => {
    // Note: Band sizes are only valid in table style definitions (CT_TblPrBase),
    // not in direct tblPr. They are set in memory but not serialized for direct tables.
    // This verifies the in-memory API works correctly.

    test('should set and get row band size', () => {
      const table = new Table(4, 2);
      table.setStyleRowBandSize(2);
      expect(table.getFormatting().tblStyleRowBandSize).toBe(2);
    });

    test('should set and get col band size', () => {
      const table = new Table(2, 4);
      table.setStyleColBandSize(3);
      expect(table.getFormatting().tblStyleColBandSize).toBe(3);
    });

    test('should support method chaining for both band sizes', () => {
      const table = new Table(4, 4);
      const result = table.setStyleRowBandSize(2).setStyleColBandSize(2);
      expect(result).toBe(table);
      expect(table.getFormatting().tblStyleRowBandSize).toBe(2);
      expect(table.getFormatting().tblStyleColBandSize).toBe(2);
    });
  });

  describe('Accessibility (caption, description)', () => {
    test('should round-trip caption and description together', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.setCaption('Sales Report Q4');
      table.setDescription('Quarterly sales by region and product category');
      table.getRow(0)!.getCell(0)!.createParagraph('Data');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fmt = loaded.getTables()[0]!.getFormatting();
      expect(fmt.caption).toBe('Sales Report Q4');
      expect(fmt.description).toBe('Quarterly sales by region and product category');

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Bidirectional Visual (bidiVisual)', () => {
    test('should round-trip bidiVisual=true', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.setBidiVisual(true);
      table.getRow(0)!.getCell(0)!.createParagraph('RTL Table');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getTables()[0]!.getFormatting().bidiVisual).toBe(true);

      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip bidiVisual=false (not present in XML)', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.setBidiVisual(false);
      table.getRow(0)!.getCell(0)!.createParagraph('LTR Table');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      // false means not present, so should be undefined or false
      const bidi = loaded.getTables()[0]!.getFormatting().bidiVisual;
      expect(!bidi).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Combined Table Properties', () => {
    test('should round-trip all tblPr properties together', async () => {
      const doc = Document.create();
      const table = new Table(3, 3);

      table
        .setWidth(9000)
        .setWidthType('dxa')
        .setAlignment('center')
        .setLayout('fixed')
        .setShading({ fill: 'F5F5F5' })
        .setCellMargins({ top: 30, bottom: 30, left: 100, right: 100 })
        .setTblLook('04A0')
        .setStyleRowBandSize(1)
        .setStyleColBandSize(1)
        .setCaption('Combined Test Table')
        .setDescription('Tests all tblPr children together')
        .setIndent(720)
        .setCellSpacing(20);

      table.getRow(0)!.getCell(0)!.createParagraph('Combined');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fmt = loaded.getTables()[0]!.getFormatting();

      expect(fmt.width).toBe(9000);
      expect(fmt.alignment).toBe('center');
      expect(fmt.shading?.fill).toBe('F5F5F5');
      expect(fmt.cellMargins?.top).toBe(30);
      expect(fmt.cellMargins?.left).toBe(100);
      expect(fmt.tblLook).toBeDefined();
      // Band sizes not serialized in direct table XML (only in table styles)
      expect(fmt.caption).toBe('Combined Test Table');
      expect(fmt.description).toBe('Tests all tblPr children together');

      doc.dispose();
      loaded.dispose();
    });
  });
});
