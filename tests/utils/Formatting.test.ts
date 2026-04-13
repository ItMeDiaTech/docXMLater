import {
  mergeFormatting,
  cloneFormatting,
  hasFormatting,
  cleanFormatting,
  isEqualFormatting,
  applyDefaults,
} from '../../src/utils/formatting';

describe('mergeFormatting', () => {
  it('should merge flat objects with override taking precedence', () => {
    const base = { bold: true, fontSize: 24 };
    const override = { fontSize: 28, italic: true };
    const result = mergeFormatting(base, override);
    expect(result).toEqual({ bold: true, fontSize: 28, italic: true });
  });

  it('should deep merge nested objects', () => {
    const base: Record<string, any> = { font: { ascii: 'Arial', eastAsia: 'SimSun' } };
    const override = { font: { ascii: 'Calibri' } };
    const result = mergeFormatting(base, override);
    expect(result.font.ascii).toBe('Calibri');
    expect(result.font.eastAsia).toBe('SimSun');
  });

  it('should skip undefined values in override', () => {
    const base = { bold: true, italic: false };
    const result = mergeFormatting(base, { bold: undefined });
    expect(result.bold).toBe(true);
  });
});

describe('cloneFormatting', () => {
  it('should create independent copy', () => {
    const original = { bold: true, indent: { left: 100 } };
    const cloned = cloneFormatting(original);
    cloned.indent.left = 200;
    expect(original.indent.left).toBe(100);
  });
});

describe('hasFormatting', () => {
  it('should return false for empty object', () => {
    expect(hasFormatting({})).toBe(false);
  });

  it('should return true when properties are defined', () => {
    expect(hasFormatting({ bold: true })).toBe(true);
    expect(hasFormatting({ fontSize: 0 })).toBe(true);
    expect(hasFormatting({ color: '' })).toBe(true);
  });

  it('should return false when all properties are undefined or null', () => {
    expect(hasFormatting({ bold: undefined, italic: null })).toBe(false);
  });
});

describe('cleanFormatting', () => {
  it('should remove undefined and null properties', () => {
    const dirty = { bold: true, italic: undefined, fontSize: null, underline: false };
    const clean = cleanFormatting(dirty);
    expect(clean).toEqual({ bold: true, underline: false });
  });

  it('should recursively clean nested objects', () => {
    const dirty = { font: { ascii: 'Arial', eastAsia: undefined }, bold: true };
    const clean = cleanFormatting(dirty);
    expect(clean).toEqual({ font: { ascii: 'Arial' }, bold: true });
  });

  it('should remove nested objects that become empty', () => {
    const dirty = { font: { ascii: undefined }, bold: true };
    const clean = cleanFormatting(dirty);
    expect(clean).toEqual({ bold: true });
  });
});

describe('isEqualFormatting', () => {
  it('should return true for identical objects', () => {
    const a = { bold: true, fontSize: 24 };
    expect(isEqualFormatting(a, { bold: true, fontSize: 24 })).toBe(true);
  });

  it('should return true for same reference', () => {
    const a = { bold: true };
    expect(isEqualFormatting(a, a)).toBe(true);
  });

  it('should return false for different values', () => {
    expect(isEqualFormatting({ bold: true }, { bold: false })).toBe(false);
  });

  it('should return false for different key counts', () => {
    expect(isEqualFormatting({ bold: true }, { bold: true, italic: true })).toBe(false);
  });

  it('should handle nested object comparison', () => {
    const a = { font: { ascii: 'Arial' } };
    const b = { font: { ascii: 'Arial' } };
    const c = { font: { ascii: 'Calibri' } };
    expect(isEqualFormatting(a, b)).toBe(true);
    expect(isEqualFormatting(a, c)).toBe(false);
  });

  it('should treat null and undefined as equal', () => {
    expect(isEqualFormatting({ a: null }, { a: undefined })).toBe(true);
  });
});

describe('applyDefaults', () => {
  it('should fill in missing properties from defaults', () => {
    const format = { bold: true };
    const defaults = { bold: false, italic: false, fontSize: 24 };
    const result = applyDefaults(format, defaults);
    expect(result).toEqual({ bold: true, italic: false, fontSize: 24 });
  });

  it('should not override defined values', () => {
    const result = applyDefaults({ fontSize: 28 }, { fontSize: 24, bold: false });
    expect(result.fontSize).toBe(28);
  });

  it('should deep merge nested defaults', () => {
    const format: Record<string, any> = { font: { ascii: 'Calibri' } };
    const defaults: Record<string, any> = {
      font: { ascii: 'Arial', eastAsia: 'SimSun' },
      bold: false,
    };
    const result = applyDefaults(format, defaults);
    expect(result.font.ascii).toBe('Calibri');
    expect(result.font.eastAsia).toBe('SimSun');
  });
});
