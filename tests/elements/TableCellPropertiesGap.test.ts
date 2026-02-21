/**
 * Gap Tests for Table Cell Properties
 *
 * Tests implemented-but-undertested TableCell.ts features:
 * - setNoWrap()
 * - setHideMark()
 * - setFitText()
 * - setHorizontalMerge() (legacy hMerge)
 */

import { Table } from '../../src/elements/Table';
import { Document } from '../../src/core/Document';

describe('Table Cell Properties Gap Tests', () => {
  describe('No Wrap (w:noWrap)', () => {
    test('should round-trip noWrap=true', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getRow(0)!.getCell(0)!.setNoWrap(true);
      table.getRow(0)!.getCell(0)!.createParagraph('No wrapping text');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getTables()[0]!.getRow(0)!.getCell(0)!.getNoWrap()).toBe(true);

      doc.dispose();
      loaded.dispose();
    });

    test('should default to false when not set', () => {
      const table = new Table(2, 2);
      expect(table.getRow(0)!.getCell(0)!.getNoWrap()).toBe(false);
    });
  });

  describe('Hide Mark (w:hideMark)', () => {
    test('should round-trip hideMark=true', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getRow(0)!.getCell(0)!.setHideMark(true);
      table.getRow(0)!.getCell(0)!.createParagraph('Hidden mark');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getTables()[0]!.getRow(0)!.getCell(0)!.getHideMark()).toBe(true);

      doc.dispose();
      loaded.dispose();
    });
  });

  describe('Fit Text (w:tcFitText)', () => {
    test('should round-trip fitText=true', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getRow(0)!.getCell(0)!.setFitText(true);
      table.getRow(0)!.getCell(0)!.createParagraph('Fit to cell');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      expect(loaded.getTables()[0]!.getRow(0)!.getCell(0)!.getFitText()).toBe(true);

      doc.dispose();
      loaded.dispose();
    });

    test('should default to false when not set', () => {
      const table = new Table(2, 2);
      expect(table.getRow(0)!.getCell(0)!.getFitText()).toBe(false);
    });
  });

  describe('Horizontal Merge (w:hMerge - legacy)', () => {
    test('should round-trip hMerge restart', async () => {
      const doc = Document.create();
      const table = new Table(2, 3);
      table.getRow(0)!.getCell(0)!.setHorizontalMerge('restart');
      table.getRow(0)!.getCell(1)!.setHorizontalMerge('continue');
      table.getRow(0)!.getCell(0)!.createParagraph('Merged');
      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const row = loaded.getTables()[0]!.getRow(0)!;
      expect(row.getCell(0)!.getHorizontalMerge()).toBe('restart');
      expect(row.getCell(1)!.getHorizontalMerge()).toBe('continue');

      doc.dispose();
      loaded.dispose();
    });

    test('should return undefined when no hMerge', () => {
      const table = new Table(2, 2);
      expect(table.getRow(0)!.getCell(0)!.getHorizontalMerge()).toBeUndefined();
    });
  });

  describe('Combined Cell Properties', () => {
    test('should round-trip multiple cell properties together', async () => {
      const doc = Document.create();
      const table = new Table(2, 3);

      const cell = table.getRow(0)!.getCell(0)!;
      cell.setNoWrap(true).setHideMark(true).setFitText(true);
      cell.setWidth(2880).setVerticalAlignment('center');
      cell.setShading({ fill: 'FFFF00' });
      cell.createParagraph('Combined');

      doc.addTable(table);

      const buffer = await doc.toBuffer();
      const loaded = await Document.loadFromBuffer(buffer);
      const loadedCell = loaded.getTables()[0]!.getRow(0)!.getCell(0)!;

      expect(loadedCell.getNoWrap()).toBe(true);
      expect(loadedCell.getHideMark()).toBe(true);
      expect(loadedCell.getFitText()).toBe(true);
      expect(loadedCell.getFormatting().width).toBe(2880);
      expect(loadedCell.getFormatting().verticalAlignment).toBe('center');
      expect(loadedCell.getFormatting().shading?.fill).toBe('FFFF00');

      doc.dispose();
      loaded.dispose();
    });
  });
});
