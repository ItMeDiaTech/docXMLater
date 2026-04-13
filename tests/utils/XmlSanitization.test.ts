import {
  removeInvalidXmlChars,
  findInvalidXmlChars,
  hasInvalidXmlChars,
  XML_CONTROL_CHARS,
} from '../../src/utils/xmlSanitization';

describe('removeInvalidXmlChars', () => {
  it('should remove NULL bytes', () => {
    expect(removeInvalidXmlChars('Hello\x00World', false)).toBe('HelloWorld');
  });

  it('should remove backspace characters', () => {
    expect(removeInvalidXmlChars('Hello\x08World', false)).toBe('HelloWorld');
  });

  it('should remove DELETE character', () => {
    expect(removeInvalidXmlChars('Hello\x7FWorld', false)).toBe('HelloWorld');
  });

  it('should preserve tab characters', () => {
    expect(removeInvalidXmlChars('Hello\tWorld', false)).toBe('Hello\tWorld');
  });

  it('should preserve newline characters', () => {
    expect(removeInvalidXmlChars('Hello\nWorld', false)).toBe('Hello\nWorld');
  });

  it('should preserve carriage return', () => {
    expect(removeInvalidXmlChars('Hello\rWorld', false)).toBe('Hello\rWorld');
  });

  it('should remove multiple invalid chars', () => {
    expect(removeInvalidXmlChars('\x00\x01\x02Hello\x1FWorld\x7F', false)).toBe('HelloWorld');
  });

  it('should return clean text unchanged', () => {
    const text = 'Normal text with spaces and punctuation! @#$%';
    expect(removeInvalidXmlChars(text, false)).toBe(text);
  });

  it('should handle empty string', () => {
    expect(removeInvalidXmlChars('', false)).toBe('');
  });
});

describe('findInvalidXmlChars', () => {
  it('should find NULL byte', () => {
    expect(findInvalidXmlChars('Hello\x00World')).toEqual([0]);
  });

  it('should find multiple different invalid chars', () => {
    const result = findInvalidXmlChars('\x00Hello\x08World\x7F');
    expect(result).toContain(0x00);
    expect(result).toContain(0x08);
    expect(result).toContain(0x7f);
  });

  it('should return unique codes only', () => {
    const result = findInvalidXmlChars('\x00\x00\x00');
    expect(result).toEqual([0]);
  });

  it('should return empty array for valid text', () => {
    expect(findInvalidXmlChars('Hello\tWorld\n')).toEqual([]);
  });

  it('should not flag tab, newline, or carriage return', () => {
    expect(findInvalidXmlChars('\t\n\r')).toEqual([]);
  });
});

describe('hasInvalidXmlChars', () => {
  it('should detect NULL byte', () => {
    expect(hasInvalidXmlChars('Hello\x00')).toBe(true);
  });

  it('should return false for clean text', () => {
    expect(hasInvalidXmlChars('Normal text')).toBe(false);
  });

  it('should return false for valid control chars', () => {
    expect(hasInvalidXmlChars('Tab\there\nNewline\rCR')).toBe(false);
  });

  it('should detect vertical tab', () => {
    expect(hasInvalidXmlChars('text\x0Bmore')).toBe(true);
  });

  it('should detect form feed', () => {
    expect(hasInvalidXmlChars('text\x0Cmore')).toBe(true);
  });
});

describe('XML_CONTROL_CHARS constants', () => {
  it('should have correct values for valid chars', () => {
    expect(XML_CONTROL_CHARS.TAB).toBe(0x09);
    expect(XML_CONTROL_CHARS.LF).toBe(0x0a);
    expect(XML_CONTROL_CHARS.CR).toBe(0x0d);
  });

  it('should have correct values for invalid chars', () => {
    expect(XML_CONTROL_CHARS.NULL).toBe(0x00);
    expect(XML_CONTROL_CHARS.BS).toBe(0x08);
    expect(XML_CONTROL_CHARS.DEL).toBe(0x7f);
    expect(XML_CONTROL_CHARS.VT).toBe(0x0b);
    expect(XML_CONTROL_CHARS.FF).toBe(0x0c);
  });
});
