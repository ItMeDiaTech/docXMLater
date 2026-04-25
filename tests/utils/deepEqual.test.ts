import { deepEqual } from '../../src/utils/deepEqual';

describe('deepEqual', () => {
  it('handles primitives identical and different', () => {
    expect(deepEqual(1, 1)).toBe(true);
    expect(deepEqual('a', 'a')).toBe(true);
    expect(deepEqual(true, true)).toBe(true);
    expect(deepEqual(0, 1)).toBe(false);
    expect(deepEqual('a', 'b')).toBe(false);
  });

  it('handles null / undefined edge cases', () => {
    expect(deepEqual(null, null)).toBe(true);
    expect(deepEqual(undefined, undefined)).toBe(true);
    expect(deepEqual(null, undefined)).toBe(false);
    expect(deepEqual(null, 0)).toBe(false);
    expect(deepEqual(undefined, {})).toBe(false);
  });

  it('compares plain objects deeply', () => {
    expect(deepEqual({ a: 1, b: 2 }, { a: 1, b: 2 })).toBe(true);
    expect(deepEqual({ a: 1, b: 2 }, { b: 2, a: 1 })).toBe(true);
    expect(deepEqual({ a: 1, b: 2 }, { a: 1, b: 3 })).toBe(false);
    expect(deepEqual({ a: 1 }, { a: 1, b: 2 })).toBe(false);
  });

  it('compares nested objects', () => {
    expect(deepEqual({ a: { b: { c: 1 } } }, { a: { b: { c: 1 } } })).toBe(true);
    expect(deepEqual({ a: { b: { c: 1 } } }, { a: { b: { c: 2 } } })).toBe(false);
  });

  it('compares arrays element-wise', () => {
    expect(deepEqual([1, 2, 3], [1, 2, 3])).toBe(true);
    expect(deepEqual([1, 2, 3], [3, 2, 1])).toBe(false);
    expect(deepEqual([1, 2], [1, 2, 3])).toBe(false);
    expect(deepEqual([{ a: 1 }], [{ a: 1 }])).toBe(true);
  });

  it('compares Date values by getTime', () => {
    expect(deepEqual(new Date(1000), new Date(1000))).toBe(true);
    expect(deepEqual(new Date(1000), new Date(2000))).toBe(false);
    expect(deepEqual(new Date(1000), 1000)).toBe(false);
  });

  it('handles array vs object asymmetry', () => {
    expect(deepEqual([], {})).toBe(false);
    expect(deepEqual([1, 2], { 0: 1, 1: 2, length: 2 })).toBe(false);
  });

  it('matches the OOXML formatting object shapes used in Paragraph/tracking', () => {
    const a = {
      spacing: { before: 100, after: 100, line: 240, lineRule: 'auto' },
      borders: { top: { style: 'single', size: 4 } },
      shading: { fill: 'FFFFFF', pattern: 'clear' },
    };
    const b = {
      borders: { top: { style: 'single', size: 4 } },
      spacing: { lineRule: 'auto', before: 100, after: 100, line: 240 },
      shading: { pattern: 'clear', fill: 'FFFFFF' },
    };
    expect(deepEqual(a, b)).toBe(true);

    const c = { ...b, spacing: { ...b.spacing, before: 101 } };
    expect(deepEqual(a, c)).toBe(false);
  });
});
