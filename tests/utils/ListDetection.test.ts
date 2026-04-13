import {
  detectTypedPrefix,
  inferLevelFromIndentation,
  inferLevelFromRelativeIndentation,
  getLevelFromFormat,
  TYPED_LIST_PATTERNS,
  PATTERN_TO_CATEGORY,
} from '../../src/utils/list-detection';

describe('detectTypedPrefix', () => {
  it('should detect decimal numbered prefixes', () => {
    const result = detectTypedPrefix('1. First item');
    expect(result.format).toBe('decimal');
    expect(result.category).toBe('numbered');
    expect(result.prefix).toBe('1. ');
  });

  it('should detect multi-digit decimal prefixes', () => {
    const result = detectTypedPrefix('12) Twelfth item');
    expect(result.format).toBe('decimal');
    expect(result.category).toBe('numbered');
  });

  it('should detect lowercase letter prefixes', () => {
    const result = detectTypedPrefix('a. Sub item');
    expect(result.format).toBe('lowerLetter');
    expect(result.category).toBe('numbered');
  });

  it('should detect uppercase letter prefixes', () => {
    const result = detectTypedPrefix('A) First section');
    expect(result.format).toBe('upperLetter');
    expect(result.category).toBe('numbered');
  });

  it('should detect bullet prefixes', () => {
    const bullet = detectTypedPrefix('• Bullet item');
    expect(bullet.format).toBe('bullet');
    expect(bullet.category).toBe('bullet');
  });

  it('should detect dash prefixes', () => {
    const dash = detectTypedPrefix('- Dash item');
    expect(dash.format).toBe('dash');
    expect(dash.category).toBe('bullet');
  });

  it('should detect arrow prefixes', () => {
    const arrow = detectTypedPrefix('► Arrow item');
    expect(arrow.format).toBe('arrow');
    expect(arrow.category).toBe('bullet');
  });

  it('should return none for plain text', () => {
    const result = detectTypedPrefix('Just normal text');
    expect(result.prefix).toBeNull();
    expect(result.format).toBeNull();
    expect(result.category).toBe('none');
  });

  it('should not match abbreviations like P.O. Box', () => {
    const result = detectTypedPrefix('P.O. Box 123');
    // Should not detect "P." as an uppercase letter list marker
    expect(result.format).not.toBe('upperLetter');
  });

  it('should not match abbreviations like U.S. Army', () => {
    const result = detectTypedPrefix('U.S. Army');
    expect(result.format).not.toBe('upperLetter');
  });
});

describe('inferLevelFromIndentation', () => {
  it('should return level 0 for standard base indent (720 twips)', () => {
    expect(inferLevelFromIndentation(720)).toBe(0);
  });

  it('should return level 1 for 1080 twips', () => {
    expect(inferLevelFromIndentation(1080)).toBe(1);
  });

  it('should return level 2 for 1440 twips', () => {
    expect(inferLevelFromIndentation(1440)).toBe(2);
  });

  it('should return level 0 for indentation below base', () => {
    expect(inferLevelFromIndentation(360)).toBe(0);
    expect(inferLevelFromIndentation(0)).toBe(0);
  });
});

describe('inferLevelFromRelativeIndentation', () => {
  it('should return level 0 for zero or negative indent', () => {
    expect(inferLevelFromRelativeIndentation(0)).toBe(0);
    expect(inferLevelFromRelativeIndentation(-100)).toBe(0);
  });

  it('should return level 1 for 360 twips', () => {
    expect(inferLevelFromRelativeIndentation(360)).toBe(1);
  });

  it('should cap at level 8', () => {
    expect(inferLevelFromRelativeIndentation(10000)).toBe(8);
  });
});

describe('getLevelFromFormat', () => {
  it('should return 0 for decimal', () => {
    expect(getLevelFromFormat('decimal')).toBe(0);
  });

  it('should return 1 for lowerLetter', () => {
    expect(getLevelFromFormat('lowerLetter')).toBe(1);
  });

  it('should return 2 for lowerRoman', () => {
    expect(getLevelFromFormat('lowerRoman')).toBe(2);
  });

  it('should return 0 for null', () => {
    expect(getLevelFromFormat(null)).toBe(0);
  });

  it('should return 0 for unknown format', () => {
    expect(getLevelFromFormat('unknown')).toBe(0);
  });
});

describe('constants', () => {
  it('should have patterns for all expected formats', () => {
    expect(TYPED_LIST_PATTERNS).toHaveProperty('decimal');
    expect(TYPED_LIST_PATTERNS).toHaveProperty('lowerLetter');
    expect(TYPED_LIST_PATTERNS).toHaveProperty('upperLetter');
    expect(TYPED_LIST_PATTERNS).toHaveProperty('lowerRoman');
    expect(TYPED_LIST_PATTERNS).toHaveProperty('bullet');
    expect(TYPED_LIST_PATTERNS).toHaveProperty('dash');
    expect(TYPED_LIST_PATTERNS).toHaveProperty('arrow');
  });

  it('should map all patterns to categories', () => {
    for (const key of Object.keys(TYPED_LIST_PATTERNS)) {
      expect(PATTERN_TO_CATEGORY[key]).toBeDefined();
    }
  });
});
