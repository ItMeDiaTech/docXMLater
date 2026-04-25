/**
 * TextBorder.style — FullBorderStyle coverage.
 *
 * Previously `TextBorderStyle` aliased `ExtendedBorderStyle` (10 values).
 * Per ECMA-376 §17.18.2 ST_Border allows 25+ values. Widened to
 * `FullBorderStyle` so consumers can assign the broader set — e.g. the
 * multi-line `thinThickSmallGap` / `thickThinMediumGap` / `triple` /
 * `dotDash` / `inset` / `outset` styles — without TypeScript rejection.
 */

import { Run } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('TextBorder.style — FullBorderStyle coverage (§17.18.2)', () => {
  const EXTENDED_VALUES = [
    'triple',
    'dotDash',
    'dotDotDash',
    'thinThickSmallGap',
    'thickThinSmallGap',
    'thinThickThinSmallGap',
    'thinThickMediumGap',
    'thickThinMediumGap',
    'thinThickThinMediumGap',
    'thinThickLargeGap',
    'thickThinLargeGap',
    'thinThickThinLargeGap',
    'dashSmallGap',
    'outset',
    'inset',
  ] as const;

  for (const style of EXTENDED_VALUES) {
    it(`accepts style: "${style}" and round-trips via toXML()`, () => {
      const run = new Run('test');
      run.setBorder({ style, size: 4, color: '000000' });
      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toContain(`w:val="${style}"`);
    });
  }
});
