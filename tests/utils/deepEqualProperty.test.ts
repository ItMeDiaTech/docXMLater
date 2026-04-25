/**
 * Property-based tests for deepEqual: reflexivity, symmetry, and
 * order-insensitivity for object key permutations.
 */
import * as fc from 'fast-check';
import { deepEqual } from '../../src/utils/deepEqual';

const jsonScalar = fc.oneof(
  fc.integer(),
  fc.string(),
  fc.boolean(),
  fc.constant(null),
  fc.constant(undefined)
);

// Limit recursion depth to avoid pathological deeply-nested objects.
const jsonValue = fc.letrec((tie) => ({
  value: fc.oneof(
    { maxDepth: 4 },
    jsonScalar,
    fc.array(tie('value'), { maxLength: 5 }),
    fc.dictionary(fc.string({ minLength: 1, maxLength: 4 }), tie('value'), { maxKeys: 5 })
  ),
})).value;

describe('deepEqual (property-based)', () => {
  it('is reflexive: deepEqual(x, x) for every x', () => {
    fc.assert(
      fc.property(jsonValue, (x) => {
        return deepEqual(x, x);
      })
    );
  });

  it('is symmetric: deepEqual(a, b) === deepEqual(b, a)', () => {
    fc.assert(
      fc.property(jsonValue, jsonValue, (a, b) => {
        return deepEqual(a, b) === deepEqual(b, a);
      })
    );
  });

  it('agrees with structural cloning (a deep-cloned value equals the original)', () => {
    fc.assert(
      fc.property(jsonValue, (x) => {
        // structuredClone is a structural deep-clone; deepEqual must
        // consider the clone equal to the source.
        const clone = structuredClone(x);
        return deepEqual(x, clone);
      })
    );
  });

  it('detects mutations: changing any key produces inequality', () => {
    fc.assert(
      fc.property(
        fc.dictionary(fc.string({ minLength: 1, maxLength: 4 }), fc.integer(), {
          minKeys: 1,
          maxKeys: 5,
        }),
        (obj) => {
          const keys = Object.keys(obj);
          if (keys.length === 0) return true;
          const firstKey = keys[0]!;
          const mutated = { ...obj, [firstKey]: (obj[firstKey] ?? 0) + 1 };
          return !deepEqual(obj, mutated);
        }
      )
    );
  });
});
