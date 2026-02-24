/**
 * Tests for text diff utility used in character-level granular tracked changes
 */

import { diffText, diffHasUnchangedParts, DiffSegment } from '../../src/utils/textDiff';

describe('textDiff', () => {
  describe('diffText', () => {
    it('should return single equal segment for identical strings', () => {
      const result = diffText('hello', 'hello');
      expect(result).toEqual([{ type: 'equal', text: 'hello' }]);
    });

    it('should return empty array for two empty strings', () => {
      const result = diffText('', '');
      expect(result).toEqual([]);
    });

    it('should return insert for empty old text', () => {
      const result = diffText('', 'new text');
      expect(result).toEqual([{ type: 'insert', text: 'new text' }]);
    });

    it('should return delete for empty new text', () => {
      const result = diffText('old text', '');
      expect(result).toEqual([{ type: 'delete', text: 'old text' }]);
    });

    it('should detect single space removal', () => {
      const result = diffText('word  word', 'word word');
      expect(result).toEqual([
        { type: 'equal', text: 'word ' },
        { type: 'delete', text: ' ' },
        { type: 'equal', text: 'word' },
      ]);
    });

    it('should detect word replacement in middle', () => {
      const result = diffText('The quick fox', 'The slow fox');
      expect(result).toEqual([
        { type: 'equal', text: 'The ' },
        { type: 'delete', text: 'quick' },
        { type: 'insert', text: 'slow' },
        { type: 'equal', text: ' fox' },
      ]);
    });

    it('should detect prefix change', () => {
      const result = diffText('Hello World', 'Goodbye World');
      expect(result).toEqual([
        { type: 'delete', text: 'Hello' },
        { type: 'insert', text: 'Goodbye' },
        { type: 'equal', text: ' World' },
      ]);
    });

    it('should detect suffix change', () => {
      const result = diffText('Hello World', 'Hello Earth');
      expect(result).toEqual([
        { type: 'equal', text: 'Hello ' },
        { type: 'delete', text: 'World' },
        { type: 'insert', text: 'Earth' },
      ]);
    });

    it('should handle complete replacement with no common text', () => {
      const result = diffText('abc', 'xyz');
      expect(result).toEqual([
        { type: 'delete', text: 'abc' },
        { type: 'insert', text: 'xyz' },
      ]);
    });

    it('should handle insertion at the end', () => {
      const result = diffText('hello', 'hello world');
      expect(result).toEqual([
        { type: 'equal', text: 'hello' },
        { type: 'insert', text: ' world' },
      ]);
    });

    it('should handle insertion at the beginning', () => {
      const result = diffText('world', 'hello world');
      expect(result).toEqual([
        { type: 'insert', text: 'hello ' },
        { type: 'equal', text: 'world' },
      ]);
    });

    it('should handle deletion at the end', () => {
      const result = diffText('hello world', 'hello');
      expect(result).toEqual([
        { type: 'equal', text: 'hello' },
        { type: 'delete', text: ' world' },
      ]);
    });

    it('should handle deletion at the beginning', () => {
      const result = diffText('hello world', 'world');
      expect(result).toEqual([
        { type: 'delete', text: 'hello ' },
        { type: 'equal', text: 'world' },
      ]);
    });

    it('should handle single character change', () => {
      const result = diffText('cat', 'bat');
      expect(result).toEqual([
        { type: 'delete', text: 'c' },
        { type: 'insert', text: 'b' },
        { type: 'equal', text: 'at' },
      ]);
    });

    it('should handle tabs and breaks in text', () => {
      const result = diffText('hello\tworld', 'hello\tearth');
      expect(result).toEqual([
        { type: 'equal', text: 'hello\t' },
        { type: 'delete', text: 'world' },
        { type: 'insert', text: 'earth' },
      ]);
    });

    it('should handle text growing longer', () => {
      const result = diffText('hi', 'hi there');
      expect(result).toEqual([
        { type: 'equal', text: 'hi' },
        { type: 'insert', text: ' there' },
      ]);
    });

    it('should handle text getting shorter', () => {
      const result = diffText('hello there', 'hello');
      expect(result).toEqual([
        { type: 'equal', text: 'hello' },
        { type: 'delete', text: ' there' },
      ]);
    });

    it('should handle replacing one char in middle of longer string', () => {
      const result = diffText('abcde', 'abXde');
      expect(result).toEqual([
        { type: 'equal', text: 'ab' },
        { type: 'delete', text: 'c' },
        { type: 'insert', text: 'X' },
        { type: 'equal', text: 'de' },
      ]);
    });

    it('should handle emoji replacement', () => {
      const result = diffText('Hello \uD83D\uDE00 World', 'Hello \uD83D\uDE01 World');
      // Common prefix: "Hello " (6 chars), common suffix: " World" (6 chars)
      // Middle changes: emoji replaced
      expect(result.length).toBeGreaterThanOrEqual(2);
      // Verify unchanged parts are preserved
      const equalParts = result.filter((s) => s.type === 'equal');
      const equalText = equalParts.map((s) => s.text).join('');
      expect(equalText).toContain('Hello ');
      expect(equalText).toContain(' World');
      // Verify the emoji region is the changed portion
      const deletePart = result.find((s) => s.type === 'delete');
      const insertPart = result.find((s) => s.type === 'insert');
      expect(deletePart).toBeDefined();
      expect(insertPart).toBeDefined();
    });

    it('should handle text with mixed emoji and ASCII', () => {
      const result = diffText('Start 🎉 end', 'Start 🚀 end');
      expect(result).toEqual([
        { type: 'equal', text: 'Start ' },
        { type: 'delete', text: '🎉' },
        { type: 'insert', text: '🚀' },
        { type: 'equal', text: ' end' },
      ]);
    });
  });

  describe('diffHasUnchangedParts', () => {
    it('should return true when there are equal segments', () => {
      const segments: DiffSegment[] = [
        { type: 'equal', text: 'hello ' },
        { type: 'delete', text: 'world' },
        { type: 'insert', text: 'earth' },
      ];
      expect(diffHasUnchangedParts(segments)).toBe(true);
    });

    it('should return false when no equal segments', () => {
      const segments: DiffSegment[] = [
        { type: 'delete', text: 'abc' },
        { type: 'insert', text: 'xyz' },
      ];
      expect(diffHasUnchangedParts(segments)).toBe(false);
    });

    it('should return false for empty segments', () => {
      expect(diffHasUnchangedParts([])).toBe(false);
    });

    it('should return true for only equal segments', () => {
      const segments: DiffSegment[] = [{ type: 'equal', text: 'hello' }];
      expect(diffHasUnchangedParts(segments)).toBe(true);
    });
  });
});
