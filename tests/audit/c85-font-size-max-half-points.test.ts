/**
 * Tests for validateFontSize upper bound
 *
 * Word's font-size limit is 1-1638 POINTS, i.e. 2-3276 half-points for w:sz
 * (ST_HpsMeasure). The validator operates in the half-point domain, so its
 * maximum must be 3276, not the point-domain constant 1638.
 */

import { validateFontSize } from '../../src/utils/validation';

describe('validateFontSize half-point range', () => {
  test('accepts the full Word-legal range up to 3276 half-points (1638 points)', () => {
    expect(() => validateFontSize(3276)).not.toThrow();
    // 820 points (w:sz 1640) — valid in Word, previously rejected
    expect(() => validateFontSize(1640)).not.toThrow();
    expect(() => validateFontSize(1638)).not.toThrow();
    expect(() => validateFontSize(2)).not.toThrow();
  });

  test('rejects sizes above 3276 half-points', () => {
    expect(() => validateFontSize(3278)).toThrow('out of range');
    expect(() => validateFontSize(3278)).toThrow('2-3276');
  });

  test('rejects sizes below 2 half-points', () => {
    expect(() => validateFontSize(0)).toThrow('out of range');
    expect(() => validateFontSize(1)).toThrow('out of range');
  });

  test('error message reports the point-domain range as 1-1638 points', () => {
    expect(() => validateFontSize(4000)).toThrow('1-1638 points');
  });

  test('still rejects non-integer and non-finite values', () => {
    expect(() => validateFontSize(24.5)).toThrow('integer');
    expect(() => validateFontSize(NaN)).toThrow('finite');
  });
});
