/**
 * Conditional formatting priority order (w:cnfStyle resolution)
 *
 * ECMA-376 §17.7.6.6 applies conditional table formats in sequence with
 * subsequent formats overriding earlier ones (banding, then first row /
 * last row, first column / last column, then corners nw, ne, sw, se).
 * First-match resolution must therefore rank the last-applied member of
 * each group first: lastRow beats firstRow, lastCol beats firstCol, and
 * seCell beats nwCell when both are co-active (single-row, single-column,
 * or 1x1 tables).
 */

import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { Style } from '../../src/formatting/Style';
import { resolveCellShading } from '../../src/processors/ShadingResolver';
import {
  getActiveConditionalsInPriorityOrder,
  CONDITIONAL_PRIORITY_ORDER,
} from '../../src/processors/cnfStyleDecoder';

describe('C69: cnfStyle priority order ranks last-applied conditional first', () => {
  describe('getActiveConditionalsInPriorityOrder', () => {
    it('ranks lastRow over firstRow when both are active (single-row table)', () => {
      const result = getActiveConditionalsInPriorityOrder('110000000000');
      expect(result).toEqual(['lastRow', 'firstRow']);
    });

    it('ranks lastCol over firstCol when both are active (single-column table)', () => {
      const result = getActiveConditionalsInPriorityOrder('001100000000');
      expect(result).toEqual(['lastCol', 'firstCol']);
    });

    it('ranks corners in reverse application order (seCell wins)', () => {
      const result = getActiveConditionalsInPriorityOrder('000000001111');
      expect(result).toEqual(['seCell', 'swCell', 'neCell', 'nwCell']);
    });

    it('ranks even banding over odd banding within each direction', () => {
      const result = getActiveConditionalsInPriorityOrder('000011110000');
      expect(result).toEqual(['band2Horz', 'band1Horz', 'band2Vert', 'band1Vert']);
    });

    it('keeps the group ranking corners > rows > columns > banding', () => {
      expect(CONDITIONAL_PRIORITY_ORDER).toEqual([
        'seCell',
        'swCell',
        'neCell',
        'nwCell',
        'lastRow',
        'firstRow',
        'lastCol',
        'firstCol',
        'band2Horz',
        'band1Horz',
        'band2Vert',
        'band1Vert',
      ]);
    });
  });

  describe('resolveCellShading with co-active pair conditionals', () => {
    it('resolves lastRow shading for a single-row table whose style defines both firstRow and lastRow', () => {
      const doc = Document.create();
      try {
        const tableStyle = new Style({
          styleId: 'PairStyle',
          name: 'Pair Style',
          type: 'table',
        });
        tableStyle.addConditionalFormatting({
          type: 'firstRow',
          cellFormatting: { shading: { fill: '111111', pattern: 'clear' } },
        });
        tableStyle.addConditionalFormatting({
          type: 'lastRow',
          cellFormatting: { shading: { fill: '222222', pattern: 'clear' } },
        });
        doc.getStylesManager().addStyle(tableStyle);

        const table = new Table(1, 2);
        table.setStyle('PairStyle');
        // Single-row table: the row is both first and last
        table.getRow(0)!.getCell(0)!.setConditionalStyle('110000000000');
        doc.addTable(table);

        const result = resolveCellShading(
          table.getRow(0)!.getCell(0)!,
          table,
          doc.getStylesManager()
        );
        expect(result).toBeDefined();
        expect(result!.fill).toBe('222222');
      } finally {
        doc.dispose();
      }
    });
  });
});
