/**
 * TableBorder.style — FullBorderStyle coverage.
 *
 * Per ECMA-376 §17.4.66 (tblBorders) and §17.18.2 ST_Border, table-level
 * borders (top / bottom / left / right / insideH / insideV) may use any
 * of 25+ spec-valid styles. `TableBorder.style` was typed as
 * `BorderStyle` (6 values only), so consumers couldn't assign spec-valid
 * multi-line-gap / triple / inset / outset styles to table borders
 * without TypeScript rejection. Widened to `FullBorderStyle` — symmetric
 * with iter 62's paragraph / cell-border widening.
 */

import { Table } from '../../src/elements/Table';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('TableBorder.style — FullBorderStyle coverage (§17.4.66 / §17.18.2)', () => {
  const VALUES = [
    'triple',
    'dotDash',
    'dotDotDash',
    'thinThickSmallGap',
    'thickThinMediumGap',
    'thinThickThinLargeGap',
    'dashSmallGap',
    'outset',
    'inset',
  ] as const;

  for (const style of VALUES) {
    it(`table outer border accepts style: "${style}"`, () => {
      const table = new Table(1, 2);
      table.setBorders({ top: { style, size: 4, color: '000000' } });
      const xml = XMLBuilder.elementToString(table.toXML());
      expect(xml).toContain(`w:val="${style}"`);
    });

    it(`table insideH border accepts style: "${style}"`, () => {
      const table = new Table(2, 2);
      table.setBorders({ insideH: { style, size: 4, color: '000000' } });
      const xml = XMLBuilder.elementToString(table.toXML());
      expect(xml).toContain(`w:val="${style}"`);
    });
  }
});
