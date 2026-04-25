/**
 * Structural equality check for OOXML formatting objects.
 *
 * Replaces `JSON.stringify(a) === JSON.stringify(b)` in hot equality paths
 * (paragraph property merging, tracking-context consolidation). Avoids
 * allocating two large strings per comparison and short-circuits on the
 * first inequality.
 *
 * Handles the value shapes seen in formatting objects: primitives, plain
 * objects, arrays, Date, null/undefined. Symbols and class instances are
 * not part of the formatting model and are compared by reference.
 */
export function deepEqual(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  // NaN === NaN is false in JS; treat structurally as equal so a parse
  // failure that produced two NaNs in equivalent slots doesn't dirty the
  // tracking-context consolidation pass.
  if (typeof a === 'number' && typeof b === 'number' && Number.isNaN(a) && Number.isNaN(b)) {
    return true;
  }
  if (a == null || b == null) return a === b;

  const typeA = typeof a;
  const typeB = typeof b;
  if (typeA !== typeB) return false;
  if (typeA !== 'object') return false;

  // Date values
  if (a instanceof Date || b instanceof Date) {
    return a instanceof Date && b instanceof Date && a.getTime() === b.getTime();
  }

  // Arrays: same length, element-wise equality.
  const aIsArr = Array.isArray(a);
  const bIsArr = Array.isArray(b);
  if (aIsArr !== bIsArr) return false;
  if (aIsArr && bIsArr) {
    const arrA = a;
    const arrB = b;
    if (arrA.length !== arrB.length) return false;
    for (let i = 0; i < arrA.length; i++) {
      if (!deepEqual(arrA[i], arrB[i])) return false;
    }
    return true;
  }

  // Plain objects: same key set, deep-equal values.
  const objA = a as Record<string, unknown>;
  const objB = b as Record<string, unknown>;
  const keysA = Object.keys(objA);
  const keysB = Object.keys(objB);
  if (keysA.length !== keysB.length) return false;
  for (const k of keysA) {
    if (!Object.prototype.hasOwnProperty.call(objB, k)) return false;
    if (!deepEqual(objA[k], objB[k])) return false;
  }
  return true;
}
