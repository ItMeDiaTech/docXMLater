/**
 * Gap Tests for Table Row Properties
 *
 * Tests implemented-but-undertested TableRow.ts features with round-trip verification:
 * - setHeight() with exact/atLeast rules
 * - setCantSplit()
 * - setHeader()
 * - setHidden()
 * - setJustification()
 * - setGridBefore() / setGridAfter()
 * - setWBefore() / setWAfter()
 * - setRowCellSpacing()
 * - setCnfStyle()
 */

import { Table } from '../../src/elements/Table';
import { Document } from '../../src/core/Document';

describe('Table Row Properties Gap Tests', () => {
  describe('Row Height (w:trHeight)', () => {
    test('should round-trip height with atLeast rule', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getRow(0)!.setHeight(720, 'atLeast');
      table.getRow(0)!.getCell(0)!.createParagraph('Tall Row');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fmt = loaded.getTables()[0]!.getRow(0)!.getFormatting();

      expect(fmt.height).toBe(720);
      expect(fmt.heightRule).toBe('atLeast');

      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip height with exact rule', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getRow(0)!.setHeight(480, 'exact');
      table.getRow(0)!.getCell(0)!.createParagraph('Exact Height');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fmt = loaded.getTables()[0]!.getRow(0)!.getFormatting();

      expect(fmt.height).toBe(480);
      expect(fmt.heightRule).toBe('exact');

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Cant Split (w:cantSplit)', () => {
    test('should round-trip cantSplit=true', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getRow(0)!.setCantSplit(true);
      table.getRow(0)!.getCell(0)!.createParagraph('No Split');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getTables()[0]!.getRow(0)!.getFormatting().cantSplit).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Header Row (w:tblHeader)', () => {
    test('should round-trip header row flag', async () => {
      const doc = Document.create();
      const table = new Table(3, 2);
      table.getRow(0)!.setHeader(true);
      table.getRow(0)!.getCell(0)!.createParagraph('Header');
      table.getRow(1)!.getCell(0)!.createParagraph('Data 1');
      table.getRow(2)!.getCell(0)!.createParagraph('Data 2');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getTables()[0]!.getRow(0)!.getFormatting().isHeader).toBe(true);
      // Non-header rows should not have header flag
      expect(!loaded.getTables()[0]!.getRow(1)!.getFormatting().isHeader).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Hidden Row (w:hidden)', () => {
    test('should round-trip hidden row', async () => {
      const doc = Document.create();
      const table = new Table(3, 2);
      table.getRow(1)!.setHidden(true);
      table.getRow(0)!.getCell(0)!.createParagraph('Visible');
      table.getRow(1)!.getCell(0)!.createParagraph('Hidden');
      table.getRow(2)!.getCell(0)!.createParagraph('Visible');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(!loaded.getTables()[0]!.getRow(0)!.getFormatting().hidden).toBe(true);
      expect(loaded.getTables()[0]!.getRow(1)!.getFormatting().hidden).toBe(true);
      expect(!loaded.getTables()[0]!.getRow(2)!.getFormatting().hidden).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Row Justification (w:jc)', () => {
    test('should round-trip all justification values', async () => {
      const values: Array<'left' | 'center' | 'right'> = ['left', 'center', 'right'];

      for (const jc of values) {
        const doc = Document.create();
        const table = new Table(2, 2);
        table.getRow(0)!.setJustification(jc);
        table.getRow(0)!.getCell(0)!.createParagraph(`Jc: ${jc}`);
        doc.addTable(table);

        const buffer = await doc.toBuffer();
        const loaded = await Document.loadFromBuffer(buffer);
        expect(loaded.getTables()[0]!.getRow(0)!.getFormatting().justification).toBe(jc);

        doc.dispose();
        loaded.dispose();
      }
    });
  });

  describe('Grid Before/After (w:gridBefore, w:gridAfter)', () => {
    test('should round-trip gridBefore', async () => {
      const doc = Document.create();
      const table = new Table(2, 4);
      table.setTableGrid([1440, 1440, 1440, 1440]);
      table.getRow(1)!.setGridBefore(2);
      table.getRow(0)!.getCell(0)!.createParagraph('Full Row');
      table.getRow(1)!.getCell(0)!.createParagraph('Offset Row');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getTables()[0]!.getRow(1)!.getFormatting().gridBefore).toBe(2);

      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip gridAfter', async () => {
      const doc = Document.create();
      const table = new Table(2, 4);
      table.setTableGrid([1440, 1440, 1440, 1440]);
      table.getRow(1)!.setGridAfter(1);
      table.getRow(0)!.getCell(0)!.createParagraph('Full Row');
      table.getRow(1)!.getCell(0)!.createParagraph('Short Row');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getTables()[0]!.getRow(1)!.getFormatting().gridAfter).toBe(1);

      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip both gridBefore and gridAfter', async () => {
      const doc = Document.create();
      const table = new Table(2, 6);
      table.setTableGrid([720, 720, 720, 720, 720, 720]);
      table.getRow(1)!.setGridBefore(1).setGridAfter(2);
      table.getRow(0)!.getCell(0)!.createParagraph('Full');
      table.getRow(1)!.getCell(0)!.createParagraph('Centered');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fmt = loaded.getTables()[0]!.getRow(1)!.getFormatting();
      expect(fmt.gridBefore).toBe(1);
      expect(fmt.gridAfter).toBe(2);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Width Before/After (w:wBefore, w:wAfter)', () => {
    test('should round-trip wBefore', async () => {
      const doc = Document.create();
      const table = new Table(2, 3);
      table.getRow(1)!.setGridBefore(1).setWBefore(1440, 'dxa');
      table.getRow(0)!.getCell(0)!.createParagraph('Full');
      table.getRow(1)!.getCell(0)!.createParagraph('Offset');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fmt = loaded.getTables()[0]!.getRow(1)!.getFormatting();
      expect(fmt.wBefore).toBe(1440);

      doc.dispose();
      loaded.dispose();
    });

    test('should round-trip wAfter', async () => {
      const doc = Document.create();
      const table = new Table(2, 3);
      table.getRow(1)!.setGridAfter(1).setWAfter(720, 'dxa');
      table.getRow(0)!.getCell(0)!.createParagraph('Full');
      table.getRow(1)!.getCell(0)!.createParagraph('Short');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fmt = loaded.getTables()[0]!.getRow(1)!.getFormatting();
      expect(fmt.wAfter).toBe(720);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Row Cell Spacing (w:tblCellSpacing on trPr)', () => {
    test('should round-trip row-level cell spacing', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getRow(0)!.setRowCellSpacing(20, 'dxa');
      table.getRow(0)!.getCell(0)!.createParagraph('Spaced');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const fmt = loaded.getTables()[0]!.getRow(0)!.getFormatting();
      expect(fmt.cellSpacing).toBe(20);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Conditional Formatting (w:cnfStyle)', () => {
    test('should round-trip cnfStyle bitmask', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getRow(0)!.setCnfStyle('100000000000'); // firstRow
      table.getRow(0)!.getCell(0)!.createParagraph('Header Row');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getTables()[0]!.getRow(0)!.getFormatting().cnfStyle).toBe('100000000000');

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Combined Row Properties', () => {
    test('should round-trip multiple row properties together', async () => {
      const doc = Document.create();
      const table = new Table(3, 3);

      const row0 = table.getRow(0)!;
      row0.setHeader(true).setCantSplit(true).setHeight(480, 'exact').setJustification('center');
      row0.getCell(0)!.createParagraph('Header');

      const row1 = table.getRow(1)!;
      row1.setCantSplit(true).setHeight(360, 'atLeast');
      row1.getCell(0)!.createParagraph('Data');

      const row2 = table.getRow(2)!;
      row2.setHidden(true);
      row2.getCell(0)!.createParagraph('Hidden');

      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const lt = loaded.getTables()[0]!;

      const fmt0 = lt.getRow(0)!.getFormatting();
      expect(fmt0.isHeader).toBe(true);
      expect(fmt0.cantSplit).toBe(true);
      expect(fmt0.height).toBe(480);
      expect(fmt0.heightRule).toBe('exact');
      expect(fmt0.justification).toBe('center');

      const fmt1 = lt.getRow(1)!.getFormatting();
      expect(fmt1.cantSplit).toBe(true);
      expect(fmt1.height).toBe(360);

      expect(lt.getRow(2)!.getFormatting().hidden).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });
});
