/**
 * XML 1.0 Char excludes U+FFFE/U+FFFF and surrogate code points outside
 * valid high/low pairs. The sanitizer previously stripped only C0 controls
 * and DEL, so a w:t containing U+FFFF produced a part Word reports as
 * unreadable, and a lone surrogate (e.g. from character-level slicing of
 * astral-plane text) was mangled at ZIP encode time.
 */
import {
  removeInvalidXmlChars,
  findInvalidXmlChars,
  hasInvalidXmlChars,
} from '../../src/utils/xmlSanitization';
import { Document } from '../../src/core/Document';

const JSZip = require('jszip');

describe('C49: noncharacters and unpaired surrogates are invalid XML 1.0', () => {
  describe('removeInvalidXmlChars', () => {
    it('strips U+FFFE and U+FFFF', () => {
      expect(removeInvalidXmlChars('a￾b￿c', false)).toBe('abc');
    });

    it('strips a lone high surrogate', () => {
      expect(removeInvalidXmlChars('a\uD800b', false)).toBe('ab');
    });

    it('strips a lone low surrogate', () => {
      expect(removeInvalidXmlChars('a\uDC00b', false)).toBe('ab');
    });

    it('preserves valid surrogate pairs (astral-plane characters)', () => {
      expect(removeInvalidXmlChars('a😀b', false)).toBe('a😀b');
    });

    it('preserves tab, newline, and carriage return alongside the new strips', () => {
      expect(removeInvalidXmlChars('a\t\n\r￿b', false)).toBe('a\t\n\rb');
    });
  });

  describe('findInvalidXmlChars', () => {
    it('reports U+FFFE and U+FFFF', () => {
      const result = findInvalidXmlChars('a￾b￿c');
      expect(result).toContain(0xfffe);
      expect(result).toContain(0xffff);
    });

    it('reports unpaired surrogates but not valid pairs', () => {
      expect(findInvalidXmlChars('a\uD800b')).toContain(0xd800);
      expect(findInvalidXmlChars('a\uDC00b')).toContain(0xdc00);
      expect(findInvalidXmlChars('a😀b')).toEqual([]);
    });
  });

  describe('hasInvalidXmlChars', () => {
    it('detects noncharacters and unpaired surrogates', () => {
      expect(hasInvalidXmlChars('a￾b')).toBe(true);
      expect(hasInvalidXmlChars('a￿b')).toBe(true);
      expect(hasInvalidXmlChars('a\uD800b')).toBe(true);
      expect(hasInvalidXmlChars('a\uDC00b')).toBe(true);
    });

    it('returns false for valid surrogate pairs', () => {
      expect(hasInvalidXmlChars('a😀b')).toBe(false);
    });
  });

  it('strips noncharacters and unpaired surrogates from saved document.xml', async () => {
    const doc = Document.create();
    doc.createParagraph('a￾b￿c\uD800d\uDC00e😀f');

    let buffer: Buffer;
    try {
      buffer = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = await JSZip.loadAsync(buffer);
    const documentXml = await zip.file('word/document.xml')!.async('string');

    // Invalid characters removed, valid astral-plane character preserved
    expect(documentXml).toContain('abcde😀f');
    expect(documentXml).not.toContain('￾');
    expect(documentXml).not.toContain('￿');
    // No replacement character from a lone surrogate hitting UTF-8 encode
    expect(documentXml).not.toContain('�');
  });
});
