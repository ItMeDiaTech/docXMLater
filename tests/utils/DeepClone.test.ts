import { deepClone } from '../../src/utils/deepClone';

describe('deepClone', () => {
  it('should clone primitive values', () => {
    expect(deepClone(42)).toBe(42);
    expect(deepClone('hello')).toBe('hello');
    expect(deepClone(true)).toBe(true);
    expect(deepClone(null)).toBeNull();
    expect(deepClone(undefined)).toBeUndefined();
  });

  it('should clone plain objects', () => {
    const original = { bold: true, color: 'FF0000', size: 24 };
    const cloned = deepClone(original);

    expect(cloned).toEqual(original);
    expect(cloned).not.toBe(original);

    cloned.bold = false;
    expect(original.bold).toBe(true);
  });

  it('should clone nested objects', () => {
    const original = {
      formatting: { font: 'Arial', style: { bold: true, italic: false } },
      spacing: { before: 120, after: 120 },
    };
    const cloned = deepClone(original);

    expect(cloned).toEqual(original);
    cloned.formatting.style.bold = false;
    expect(original.formatting.style.bold).toBe(true);
  });

  it('should clone arrays', () => {
    const original = [1, 2, [3, 4, [5]]];
    const cloned = deepClone(original);

    expect(cloned).toEqual(original);
    expect(cloned).not.toBe(original);
    expect(cloned[2]).not.toBe(original[2]);
  });

  it('should clone Date objects', () => {
    const original = new Date('2026-01-15T10:30:00Z');
    const cloned = deepClone(original);

    expect(cloned).toEqual(original);
    expect(cloned).not.toBe(original);
    expect(cloned instanceof Date).toBe(true);
    expect(cloned.getTime()).toBe(original.getTime());
  });

  it('should clone Map objects', () => {
    const original = new Map<string, number>([
      ['a', 1],
      ['b', 2],
    ]);
    const cloned = deepClone(original);

    expect(cloned.get('a')).toBe(1);
    expect(cloned.get('b')).toBe(2);
    expect(cloned).not.toBe(original);

    cloned.set('c', 3);
    expect(original.has('c')).toBe(false);
  });

  it('should clone Set objects', () => {
    const original = new Set([1, 2, 3]);
    const cloned = deepClone(original);

    expect(cloned.size).toBe(3);
    expect(cloned.has(1)).toBe(true);
    expect(cloned).not.toBe(original);

    cloned.add(4);
    expect(original.has(4)).toBe(false);
  });

  it('should clone RegExp objects', () => {
    const original = /test\d+/gi;
    const cloned = deepClone(original);

    expect(cloned.source).toBe(original.source);
    expect(cloned.flags).toBe(original.flags);
    expect(cloned).not.toBe(original);
  });

  it('should clone objects with Date properties', () => {
    const original = {
      name: 'revision',
      date: new Date('2026-03-15'),
      nested: { created: new Date('2026-01-01') },
    };
    const cloned = deepClone(original);

    expect(cloned.date instanceof Date).toBe(true);
    expect(cloned.nested.created instanceof Date).toBe(true);
    expect(cloned.date.getTime()).toBe(original.date.getTime());
  });

  it('should handle empty objects and arrays', () => {
    expect(deepClone({})).toEqual({});
    expect(deepClone([])).toEqual([]);
  });

  it('should clone formatting-like objects used in docxmlater', () => {
    const formatting = {
      bold: true,
      italic: false,
      fontSize: 24,
      color: 'FF0000',
      underline: 'single' as const,
      font: { ascii: 'Arial', eastAsia: undefined },
    };
    const cloned = deepClone(formatting);

    expect(cloned).toEqual(formatting);
    expect(cloned).not.toBe(formatting);
    expect(cloned.font).not.toBe(formatting.font);
  });
});
