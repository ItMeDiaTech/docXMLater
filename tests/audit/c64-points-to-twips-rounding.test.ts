/**
 * Tests for integer twip enforcement
 *
 * OOXML twips attributes (w:spacing w:before/w:after/w:line, w:ind) use
 * ST_TwipsMeasure/ST_SignedTwipsMeasure, which are ST_DecimalNumber-based and
 * require integer values. pointsToTwips must round (like pointsToHalfPoints,
 * pointsToEmus, inchesToTwips) and validateTwips must reject fractional input.
 */

import { pointsToTwips } from '../../src/utils/units';
import { validateTwips } from '../../src/utils/validation';

describe('pointsToTwips integer rounding', () => {
  test('rounds fractional point values to the nearest integer twip', () => {
    // 10.33 pt * 20 = 206.6 — not representable as integer twips without rounding
    expect(pointsToTwips(10.33)).toBe(207);
    expect(pointsToTwips(10.32)).toBe(206);
  });

  test('always returns an integer for arbitrary fractional inputs', () => {
    const samples = [0.1, 0.33, 1.07, 5.555, 12.345, 99.99, 10.33];
    for (const pts of samples) {
      expect(Number.isInteger(pointsToTwips(pts))).toBe(true);
    }
  });

  test('exact conversions are unchanged', () => {
    expect(pointsToTwips(0)).toBe(0);
    expect(pointsToTwips(10)).toBe(200);
    expect(pointsToTwips(10.5)).toBe(210);
    expect(pointsToTwips(-6)).toBe(-120);
  });
});

describe('validateTwips integer enforcement', () => {
  test('rejects fractional twip values', () => {
    expect(() => validateTwips(206.6)).toThrow('must be an integer');
    expect(() => validateTwips(0.5, 'spacing')).toThrow('spacing must be an integer');
    expect(() => validateTwips(-100.25)).toThrow('must be an integer');
  });

  test('accepts integer twip values within range', () => {
    expect(() => validateTwips(0)).not.toThrow();
    expect(() => validateTwips(206)).not.toThrow();
    expect(() => validateTwips(-31680)).not.toThrow();
    expect(() => validateTwips(31680)).not.toThrow();
  });

  test('still rejects non-finite and out-of-range values', () => {
    expect(() => validateTwips(NaN)).toThrow('finite');
    expect(() => validateTwips(Infinity)).toThrow('finite');
    expect(() => validateTwips(31681)).toThrow('out of range');
    expect(() => validateTwips(-31681)).toThrow('out of range');
  });
});
