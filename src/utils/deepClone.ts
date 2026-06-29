/**
 * Deep clone utility for safely cloning objects
 * More efficient and type-safe than JSON.parse(JSON.stringify())
 */

/**
 * Deep clone an object using structured cloning
 * Preserves most object types including Date, RegExp, Map, Set, etc.
 *
 * For simple objects (like formatting options), this is more efficient
 * than JSON.parse(JSON.stringify()) and doesn't lose non-serializable values.
 *
 * Cyclic and shared references are handled via an internal seen-cache: an
 * object encountered more than once during a single clone is cloned exactly
 * once, so circular graphs cannot overflow the stack and shared sub-objects
 * (including shared Date and RegExp instances) retain a single shared identity
 * in the result.
 *
 * @param obj - Object to clone
 * @returns Deep cloned copy of the object
 *
 * @example
 * ```typescript
 * const original = { bold: true, color: "FF0000", date: new Date() };
 * const cloned = deepClone(original);
 * cloned.bold = false;
 * console.log(original.bold); // true (unchanged)
 * console.log(cloned.date instanceof Date); // true (preserved)
 * ```
 */
export function deepClone<T>(obj: T): T {
  // Map of already-cloned source objects to their clones, scoped to this call.
  // Guards against infinite recursion on cyclic input and preserves the
  // identity of references shared by multiple parents within the same graph.
  return cloneInternal(obj, new WeakMap<object, unknown>());
}

function cloneInternal<T>(obj: T, seen: WeakMap<object, unknown>): T {
  // Handle primitive types and null
  if (obj === null || typeof obj !== 'object') {
    return obj;
  }

  // Return the existing clone for any object already seen in this call. Placed
  // before the Date/RegExp branches so a single Date/RegExp instance shared by
  // multiple nodes resolves to one shared clone (the JSDoc shared-identity
  // guarantee), not a fresh copy per reference.
  const existing = seen.get(obj as object);
  if (existing !== undefined) {
    return existing as T;
  }

  // Handle Date (immutable value copy; cannot form a cycle). Registered in
  // `seen` so repeat references to the same instance share this clone.
  if (obj instanceof Date) {
    const dateCopy = new Date(obj.getTime());
    seen.set(obj as object, dateCopy);
    return dateCopy as T;
  }

  // Handle RegExp (immutable value copy; cannot form a cycle). Registered in
  // `seen` so repeat references to the same instance share this clone.
  if (obj instanceof RegExp) {
    const regexpCopy = new RegExp(obj.source, obj.flags);
    seen.set(obj as object, regexpCopy);
    return regexpCopy as T;
  }

  // Handle Array
  if (Array.isArray(obj)) {
    const arrCopy: unknown[] = [];
    // Register before recursing so self/back references resolve to this clone.
    seen.set(obj as object, arrCopy);
    for (let i = 0; i < obj.length; i++) {
      arrCopy[i] = cloneInternal(obj[i], seen);
    }
    return arrCopy as T;
  }

  // Handle Map
  if (obj instanceof Map) {
    const mapCopy = new Map();
    seen.set(obj, mapCopy);
    obj.forEach((value, key) => {
      mapCopy.set(cloneInternal(key, seen), cloneInternal(value, seen));
    });
    return mapCopy as T;
  }

  // Handle Set
  if (obj instanceof Set) {
    const setCopy = new Set();
    seen.set(obj, setCopy);
    obj.forEach((value) => {
      setCopy.add(cloneInternal(value, seen));
    });
    return setCopy as T;
  }

  // Handle plain objects
  const objCopy = Object.create(Object.getPrototypeOf(obj)) as Record<string, unknown>;
  seen.set(obj as object, objCopy);
  // Object.keys avoids the prototype-builtins shadowing risk of
  // obj.hasOwnProperty(key) (matches sibling deepEqual.ts) and yields the same
  // own-enumerable string keys the previous for..in + hasOwnProperty loop did.
  for (const key of Object.keys(obj as object)) {
    objCopy[key] = cloneInternal((obj as Record<string, unknown>)[key], seen);
  }

  return objCopy as T;
}
