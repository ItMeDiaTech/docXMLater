/**
 * Property-based tests for unit conversions in src/utils/units.ts.
 *
 * Verifies algebraic round-trip identities (twips ↔ EMUs, twips ↔ inches,
 * EMUs ↔ inches) at scale rather than at hand-picked sample values.
 * fast-check generates random inputs in plausible OOXML ranges and shrinks
 * to a minimal failing case when an identity breaks.
 */
import * as fc from 'fast-check';
import {
  twipsToPoints,
  pointsToTwips,
  twipsToInches,
  inchesToTwips,
  twipsToEmus,
  emusToTwips,
  inchesToEmus,
  emusToInches,
  pointsToHalfPoints,
  halfPointsToPoints,
  inchesToPoints,
  pointsToInches,
} from '../../src/utils/units';

describe('Unit conversions (property-based)', () => {
  // OOXML twips range: 0 to ~31MM (about 21500 inches — well past any real document).
  // Cap at 1MM twips (~700 inches) for sanity.
  const finiteTwips = fc.integer({ min: 0, max: 1_000_000 });
  const finitePoints = fc.integer({ min: 0, max: 50_000 });
  const finiteInches = fc.integer({ min: 0, max: 700 });

  it('twips ↔ points round-trips for any non-negative integer twip count divisible by 20', () => {
    fc.assert(
      fc.property(finiteTwips, (twips) => {
        const aligned = twips - (twips % 20); // points = twips/20, so only multiples round-trip
        const pts = twipsToPoints(aligned);
        const back = pointsToTwips(pts);
        return back === aligned;
      })
    );
  });

  it('points ↔ twips round-trips for any non-negative integer point count', () => {
    fc.assert(
      fc.property(finitePoints, (pts) => {
        const twips = pointsToTwips(pts);
        const back = twipsToPoints(twips);
        return back === pts;
      })
    );
  });

  it('inches ↔ twips round-trips for any non-negative integer inch count', () => {
    fc.assert(
      fc.property(finiteInches, (inches) => {
        const twips = inchesToTwips(inches);
        const back = twipsToInches(twips);
        return back === inches;
      })
    );
  });

  it('twips ↔ EMUs round-trips for any non-negative integer twip count', () => {
    fc.assert(
      fc.property(finiteTwips, (twips) => {
        const emus = twipsToEmus(twips);
        const back = emusToTwips(emus);
        return back === twips;
      })
    );
  });

  it('inches ↔ EMUs round-trips for any non-negative integer inch count', () => {
    fc.assert(
      fc.property(finiteInches, (inches) => {
        const emus = inchesToEmus(inches);
        const back = emusToInches(emus);
        return back === inches;
      })
    );
  });

  it('points ↔ half-points round-trips for any non-negative integer point count', () => {
    fc.assert(
      fc.property(finitePoints, (pts) => {
        const hp = pointsToHalfPoints(pts);
        const back = halfPointsToPoints(hp);
        return back === pts;
      })
    );
  });

  it('inches → points → inches preserves integer inch counts', () => {
    fc.assert(
      fc.property(finiteInches, (inches) => {
        const pts = inchesToPoints(inches);
        const back = pointsToInches(pts);
        return back === inches;
      })
    );
  });

  it('twips conversions never produce NaN or Infinity for non-negative inputs', () => {
    fc.assert(
      fc.property(finiteTwips, (twips) => {
        return [twipsToPoints(twips), twipsToInches(twips), twipsToEmus(twips)].every((v) =>
          Number.isFinite(v)
        );
      })
    );
  });
});
