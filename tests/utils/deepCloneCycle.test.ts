import { deepClone } from '../../src/utils/deepClone';

describe('deepClone cycle and shared-reference handling', () => {
  it('clones a self-referential object without throwing and rewires the cycle', () => {
    const obj: { name: string; self?: unknown } = { name: 'root' };
    obj.self = obj;

    let cloned!: typeof obj;
    expect(() => {
      cloned = deepClone(obj);
    }).not.toThrow();

    expect(cloned).not.toBe(obj);
    expect(cloned.name).toBe('root');
    // The cycle is preserved but points at the clone, not the original.
    expect(cloned.self).toBe(cloned);
    expect(cloned.self).not.toBe(obj);
  });

  it('clones a cyclic array without throwing', () => {
    const arr: unknown[] = [1, 2];
    arr.push(arr);

    let cloned!: unknown[];
    expect(() => {
      cloned = deepClone(arr);
    }).not.toThrow();

    expect(cloned).not.toBe(arr);
    expect(cloned[0]).toBe(1);
    expect(cloned[1]).toBe(2);
    expect(cloned[2]).toBe(cloned);
  });

  it('handles indirect cycles across nested containers', () => {
    const parent: { child?: unknown } = {};
    const child: { parent?: unknown } = {};
    parent.child = child;
    child.parent = parent;

    let cloned!: typeof parent;
    expect(() => {
      cloned = deepClone(parent);
    }).not.toThrow();

    const clonedChild = cloned.child as { parent?: unknown };
    expect(clonedChild).not.toBe(child);
    expect(clonedChild.parent).toBe(cloned);
  });

  it('preserves shared (non-cyclic) references as a single cloned instance', () => {
    const shared = { value: 42 };
    const original = { left: shared, right: shared };

    const cloned = deepClone(original);

    expect(cloned.left).not.toBe(shared);
    expect(cloned.left).toEqual(shared);
    // Both fields must resolve to the SAME cloned instance, mirroring the input.
    expect(cloned.left).toBe(cloned.right);

    cloned.left.value = 99;
    expect(cloned.right.value).toBe(99);
    expect(shared.value).toBe(42);
  });

  it('preserves shared references held inside arrays, Maps, and Sets', () => {
    const shared = { tag: 'x' };
    const original = {
      list: [shared, shared],
      map: new Map<string, typeof shared>([['a', shared]]),
      set: new Set<typeof shared>([shared]),
    };

    const cloned = deepClone(original);
    const first = cloned.list[0]!;

    expect(first).toBe(cloned.list[1]);
    expect(cloned.map.get('a')).toBe(first);
    expect(cloned.set.has(first)).toBe(true);
    expect(first).not.toBe(shared);
  });

  it('still deep-clones acyclic Date, RegExp, Map, Set, and nested objects independently', () => {
    const original = {
      date: new Date('2026-03-15T08:00:00Z'),
      pattern: /abc\d+/gi,
      map: new Map<string, { n: number }>([['k', { n: 1 }]]),
      set: new Set<number>([1, 2, 3]),
      nested: { inner: { flag: true } },
    };

    const cloned = deepClone(original);

    expect(cloned).toEqual(original);
    expect(cloned).not.toBe(original);

    expect(cloned.date instanceof Date).toBe(true);
    expect(cloned.date).not.toBe(original.date);
    expect(cloned.date.getTime()).toBe(original.date.getTime());

    expect(cloned.pattern instanceof RegExp).toBe(true);
    expect(cloned.pattern.source).toBe(original.pattern.source);
    expect(cloned.pattern.flags).toBe(original.pattern.flags);
    expect(cloned.pattern).not.toBe(original.pattern);

    expect(cloned.map).not.toBe(original.map);
    expect(cloned.map.get('k')).not.toBe(original.map.get('k'));
    expect(cloned.map.get('k')).toEqual({ n: 1 });

    expect(cloned.set).not.toBe(original.set);
    expect([...cloned.set]).toEqual([1, 2, 3]);

    // Mutating the clone must not leak into the original graph.
    cloned.nested.inner.flag = false;
    cloned.map.get('k')!.n = 999;
    expect(original.nested.inner.flag).toBe(true);
    expect(original.map.get('k')!.n).toBe(1);
  });

  it('clones a Date shared by two fields to the SAME cloned instance', () => {
    const shared = new Date('2026-03-15T08:00:00Z');
    const original = { created: shared, modified: shared };

    const cloned = deepClone(original);

    expect(cloned.created instanceof Date).toBe(true);
    expect(cloned.created).not.toBe(shared);
    expect(cloned.created.getTime()).toBe(shared.getTime());
    // The single shared Date instance must yield a single shared clone.
    expect(cloned.created).toBe(cloned.modified);
  });

  it('clones a RegExp shared by two fields to the SAME cloned instance', () => {
    const shared = /abc\d+/gi;
    const original = { include: shared, exclude: shared };

    const cloned = deepClone(original);

    expect(cloned.include instanceof RegExp).toBe(true);
    expect(cloned.include).not.toBe(shared);
    expect(cloned.include.source).toBe(shared.source);
    expect(cloned.include.flags).toBe(shared.flags);
    // The single shared RegExp instance must yield a single shared clone.
    expect(cloned.include).toBe(cloned.exclude);
  });

  it('deep-clones a RunFormatting-like object with a nested object field independently', () => {
    // Mirrors the run-formatting shape passed through deepClone in the codebase.
    const original = {
      bold: true,
      italic: false,
      fontSize: 24,
      color: 'FF0000',
      underline: 'single' as const,
      font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'SimSun' },
    };

    const cloned = deepClone(original);

    expect(cloned).toEqual(original);
    expect(cloned).not.toBe(original);
    // The nested font object must be an independent copy.
    expect(cloned.font).not.toBe(original.font);

    cloned.font.ascii = 'Calibri';
    expect(original.font.ascii).toBe('Arial');
  });

  it('clones a null-prototype object without throwing', () => {
    const original = Object.create(null) as Record<string, unknown>;
    original.a = 1;
    original.nested = { b: 2 };

    let cloned!: Record<string, unknown>;
    expect(() => {
      cloned = deepClone(original);
    }).not.toThrow();

    expect(cloned).not.toBe(original);
    expect(Object.getPrototypeOf(cloned)).toBeNull();
    expect(cloned.a).toBe(1);
    expect(cloned.nested).not.toBe(original.nested);
    expect(cloned.nested).toEqual({ b: 2 });
  });
});
