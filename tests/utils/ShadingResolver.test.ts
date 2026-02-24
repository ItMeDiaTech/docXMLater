/**
 * ShadingResolver - Tests for shading inheritance resolution and bitmask decoding
 */

import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { Style } from '../../src/formatting/Style';
import { resolveCellShading } from '../../src/utils/ShadingResolver';
import {
  decodeCnfStyle,
  getActiveConditionalsInPriorityOrder,
} from '../../src/utils/cnfStyleDecoder';

describe('ShadingResolver', () => {
  describe('resolveCellShading', () => {
    it('should return direct cell shading when set', () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getRow(0)!.getCell(0)!.setShading({ fill: 'FF0000', pattern: 'solid' });
      table.setShading({ fill: '0000FF', pattern: 'solid' });
      doc.addTable(table);

      const result = resolveCellShading(
        table.getRow(0)!.getCell(0)!,
        table,
        doc.getStylesManager()
      );
      expect(result).toBeDefined();
      expect(result!.fill).toBe('FF0000');
      doc.dispose();
    });

    it('should return nil as undefined (explicitly clear)', () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getRow(0)!.getCell(0)!.setShading({ pattern: 'nil' });
      table.setShading({ fill: '0000FF', pattern: 'solid' });
      doc.addTable(table);

      const result = resolveCellShading(
        table.getRow(0)!.getCell(0)!,
        table,
        doc.getStylesManager()
      );
      expect(result).toBeUndefined();
      doc.dispose();
    });

    it('should fall back to table shading when cell has none', () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.setShading({ fill: '00FF00', pattern: 'clear' });
      doc.addTable(table);

      const result = resolveCellShading(
        table.getRow(0)!.getCell(0)!,
        table,
        doc.getStylesManager()
      );
      expect(result).toBeDefined();
      expect(result!.fill).toBe('00FF00');
      doc.dispose();
    });

    it('should return undefined when no shading exists at any level', () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      doc.addTable(table);

      const result = resolveCellShading(
        table.getRow(0)!.getCell(0)!,
        table,
        doc.getStylesManager()
      );
      expect(result).toBeUndefined();
      doc.dispose();
    });

    it('should use conditional table style shading when cnfStyle matches', () => {
      const doc = Document.create();

      // Create a table style with firstRow conditional
      const tableStyle = new Style({
        styleId: 'TestTableStyle',
        name: 'Test Table Style',
        type: 'table',
      });
      tableStyle.addConditionalFormatting({
        type: 'firstRow',
        cellFormatting: {
          shading: { fill: '4472C4', pattern: 'clear', themeFill: 'accent1' },
        },
      });
      doc.getStylesManager().addStyle(tableStyle);

      const table = new Table(2, 2);
      table.setStyle('TestTableStyle');
      // Set cnfStyle for first row cell (firstRow = bit 0)
      table.getRow(0)!.getCell(0)!.setConditionalStyle('100000000000');
      doc.addTable(table);

      const result = resolveCellShading(
        table.getRow(0)!.getCell(0)!,
        table,
        doc.getStylesManager()
      );
      expect(result).toBeDefined();
      expect(result!.fill).toBe('4472C4');
      expect(result!.themeFill).toBe('accent1');
      doc.dispose();
    });

    it('should use table style default cell shading as fallback', () => {
      const doc = Document.create();

      const tableStyle = new Style({
        styleId: 'FallbackStyle',
        name: 'Fallback Style',
        type: 'table',
      });
      tableStyle.setTableCellFormatting({
        shading: { fill: 'E0E0E0', pattern: 'clear' },
      });
      doc.getStylesManager().addStyle(tableStyle);

      const table = new Table(2, 2);
      table.setStyle('FallbackStyle');
      doc.addTable(table);

      const result = resolveCellShading(
        table.getRow(0)!.getCell(0)!,
        table,
        doc.getStylesManager()
      );
      expect(result).toBeDefined();
      expect(result!.fill).toBe('E0E0E0');
      doc.dispose();
    });

    it('should skip auto fill and continue resolution', () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getRow(0)!.getCell(0)!.setShading({ fill: 'auto', pattern: 'clear' });
      table.setShading({ fill: 'AABBCC', pattern: 'clear' });
      doc.addTable(table);

      const result = resolveCellShading(
        table.getRow(0)!.getCell(0)!,
        table,
        doc.getStylesManager()
      );
      expect(result).toBeDefined();
      expect(result!.fill).toBe('AABBCC');
      doc.dispose();
    });
  });

  describe('Document.getComputedCellShading', () => {
    it('should resolve shading via convenience method', () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      table.getRow(0)!.getCell(0)!.setShading({ fill: 'FF0000', pattern: 'solid' });
      doc.addTable(table);

      const result = doc.getComputedCellShading(table, 0, 0);
      expect(result).toBeDefined();
      expect(result!.fill).toBe('FF0000');
      doc.dispose();
    });

    it('should return undefined for invalid row/col', () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      doc.addTable(table);

      expect(doc.getComputedCellShading(table, 5, 0)).toBeUndefined();
      expect(doc.getComputedCellShading(table, 0, 5)).toBeUndefined();
      doc.dispose();
    });
  });

  describe('decodeCnfStyle', () => {
    it('should decode all-zeros bitmask', () => {
      const flags = decodeCnfStyle('000000000000');
      expect(flags.firstRow).toBe(false);
      expect(flags.lastRow).toBe(false);
      expect(flags.firstCol).toBe(false);
      expect(flags.lastCol).toBe(false);
    });

    it('should decode firstRow bit', () => {
      const flags = decodeCnfStyle('100000000000');
      expect(flags.firstRow).toBe(true);
      expect(flags.lastRow).toBe(false);
    });

    it('should decode corner cell bits', () => {
      const flags = decodeCnfStyle('000000001100');
      expect(flags.neCell).toBe(true);
      expect(flags.nwCell).toBe(true);
      expect(flags.seCell).toBe(false);
    });

    it('should decode band bits', () => {
      const flags = decodeCnfStyle('000010100000');
      expect(flags.band1Horz).toBe(true);
      expect(flags.band1Vert).toBe(true);
    });
  });

  describe('getActiveConditionalsInPriorityOrder', () => {
    it('should return corner cells before edges', () => {
      // neCell (bit 8) + firstRow (bit 0) + firstCol (bit 2)
      const result = getActiveConditionalsInPriorityOrder('101000001000');
      // neCell should be first (corners before edges)
      expect(result.indexOf('neCell')).toBeLessThan(result.indexOf('firstRow'));
      expect(result.indexOf('neCell')).toBeLessThan(result.indexOf('firstCol'));
    });

    it('should return empty for all-zero bitmask', () => {
      const result = getActiveConditionalsInPriorityOrder('000000000000');
      expect(result).toHaveLength(0);
    });
  });
});
