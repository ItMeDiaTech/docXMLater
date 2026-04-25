/**
 * BorderDefinition.style — FullBorderStyle coverage.
 *
 * `BorderDefinition` is the shared border interface used for paragraph
 * borders, table borders, and cell borders. Its `style` field was
 * previously typed as `BorderStyle | ExtendedBorderStyle` (10 values)
 * even though the generators pass the raw string straight through to
 * `<w:val="..."/>` and ECMA-376 §17.18.2 ST_Border defines 25+ values.
 * Consumers constructing paragraph or cell borders couldn't assign the
 * multi-line-gap / triple / inset / outset / etc. variants without
 * TypeScript rejection.
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { TableCell } from '../../src/elements/TableCell';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('BorderDefinition.style — FullBorderStyle coverage (§17.18.2)', () => {
  const EXTENDED_VALUES = [
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

  for (const style of EXTENDED_VALUES) {
    it(`paragraph border accepts style: "${style}"`, () => {
      const p = new Paragraph();
      p.addText('x');
      p.setBorder({ top: { style, size: 4, color: '000000' } });
      const xml = XMLBuilder.elementToString(p.toXML());
      expect(xml).toContain(`w:val="${style}"`);
    });

    it(`cell border accepts style: "${style}"`, () => {
      const cell = new TableCell();
      cell.setBorders({ top: { style, size: 4, color: '000000' } });
      const xml = XMLBuilder.elementToString(cell.toXML());
      expect(xml).toContain(`w:val="${style}"`);
    });
  }
});
