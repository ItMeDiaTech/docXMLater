/**
 * Tests for validation utilities
 */

import {
  validateDocxStructure,
  isBinaryFile,
  normalizePath,
  isValidZipBuffer,
  isTextContent,
  sanitizeHyperlinkUrl,
} from '../../src/utils/validation';
import { MissingRequiredFileError } from '../../src/zip/errors';
import { REQUIRED_DOCX_FILES } from '../../src/zip/types';

describe('Validation Utilities', () => {
  describe('validateDocxStructure', () => {
    test('should not throw for valid DOCX structure', () => {
      const validFiles = [
        '[Content_Types].xml',
        '_rels/.rels',
        'word/document.xml',
        'word/styles.xml',
      ];

      expect(() => validateDocxStructure(validFiles)).not.toThrow();
    });

    test('should throw when missing [Content_Types].xml', () => {
      const invalidFiles = ['_rels/.rels', 'word/document.xml'];

      expect(() => validateDocxStructure(invalidFiles)).toThrow(MissingRequiredFileError);
    });

    test('should throw when missing _rels/.rels', () => {
      const invalidFiles = ['[Content_Types].xml', 'word/document.xml'];

      expect(() => validateDocxStructure(invalidFiles)).toThrow(MissingRequiredFileError);
    });

    test('should throw when missing word/document.xml', () => {
      const invalidFiles = ['[Content_Types].xml', '_rels/.rels'];

      expect(() => validateDocxStructure(invalidFiles)).toThrow(MissingRequiredFileError);
    });

    test('should validate all required files', () => {
      const allRequired = Array.from(REQUIRED_DOCX_FILES);
      expect(() => validateDocxStructure(allRequired)).not.toThrow();
    });
  });

  describe('isBinaryFile', () => {
    test('should identify image files as binary', () => {
      expect(isBinaryFile('image.png')).toBe(true);
      expect(isBinaryFile('photo.jpg')).toBe(true);
      expect(isBinaryFile('picture.jpeg')).toBe(true);
      expect(isBinaryFile('icon.gif')).toBe(true);
      expect(isBinaryFile('bitmap.bmp')).toBe(true);
    });

    test('should identify font files as binary', () => {
      expect(isBinaryFile('font.ttf')).toBe(true);
      expect(isBinaryFile('font.otf')).toBe(true);
      expect(isBinaryFile('font.woff')).toBe(true);
    });

    test('should identify text files as non-binary', () => {
      expect(isBinaryFile('document.xml')).toBe(false);
      expect(isBinaryFile('readme.txt')).toBe(false);
      expect(isBinaryFile('style.css')).toBe(false);
      expect(isBinaryFile('script.js')).toBe(false);
    });

    test('should be case-insensitive', () => {
      expect(isBinaryFile('IMAGE.PNG')).toBe(true);
      expect(isBinaryFile('Photo.JPG')).toBe(true);
      expect(isBinaryFile('DOCUMENT.XML')).toBe(false);
    });

    test('should handle paths with directories', () => {
      expect(isBinaryFile('word/media/image1.png')).toBe(true);
      expect(isBinaryFile('word/document.xml')).toBe(false);
    });
  });

  describe('normalizePath', () => {
    test('should convert backslashes to forward slashes', () => {
      expect(normalizePath('path\\to\\file.txt')).toBe('path/to/file.txt');
    });

    test('should remove leading slashes', () => {
      expect(normalizePath('/path/to/file.txt')).toBe('path/to/file.txt');
      expect(normalizePath('//path/to/file.txt')).toBe('path/to/file.txt');
    });

    test('should handle mixed slashes', () => {
      expect(normalizePath('\\path/to\\file.txt')).toBe('path/to/file.txt');
    });

    test('should preserve forward slashes in middle of path', () => {
      expect(normalizePath('path/to/file.txt')).toBe('path/to/file.txt');
    });

    test('should handle empty string', () => {
      expect(normalizePath('')).toBe('');
    });

    test('should handle single file name', () => {
      expect(normalizePath('file.txt')).toBe('file.txt');
    });

    // Security tests
    test('should reject path traversal with ../ (Unix)', () => {
      expect(() => normalizePath('../etc/passwd')).toThrow('path traversal');
      expect(() => normalizePath('word/../../etc/passwd')).toThrow('path traversal');
      expect(() => normalizePath('word/media/../../../tmp/evil.sh')).toThrow('path traversal');
    });

    test('should reject path traversal with ..\\ (Windows)', () => {
      expect(() => normalizePath('..\\Windows\\System32')).toThrow('path traversal');
      expect(() => normalizePath('word\\..\\..\\tmp')).toThrow('path traversal');
    });

    test('should reject absolute Windows paths', () => {
      expect(() => normalizePath('C:\\Windows\\System32')).toThrow('absolute Windows path');
      expect(() => normalizePath('D:\\tmp\\evil.exe')).toThrow('absolute Windows path');
      expect(() => normalizePath('c:/windows/system32')).toThrow('absolute Windows path');
    });

    test('should allow valid relative paths', () => {
      expect(normalizePath('word/document.xml')).toBe('word/document.xml');
      expect(normalizePath('word/media/image1.png')).toBe('word/media/image1.png');
      expect(normalizePath('[Content_Types].xml')).toBe('[Content_Types].xml');
    });
  });

  describe('isValidZipBuffer', () => {
    test('should recognize valid ZIP signature', () => {
      const validZip = Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x00, 0x00]);
      expect(isValidZipBuffer(validZip)).toBe(true);
    });

    test('should recognize empty ZIP signature', () => {
      const emptyZip = Buffer.from([0x50, 0x4b, 0x05, 0x06, 0x00, 0x00]);
      expect(isValidZipBuffer(emptyZip)).toBe(true);
    });

    test('should reject invalid signature', () => {
      const invalid = Buffer.from([0x00, 0x00, 0x00, 0x00]);
      expect(isValidZipBuffer(invalid)).toBe(false);
    });

    test('should reject buffer too small', () => {
      const tooSmall = Buffer.from([0x50, 0x4b]);
      expect(isValidZipBuffer(tooSmall)).toBe(false);
    });

    test('should reject text files', () => {
      const textFile = Buffer.from('This is a text file');
      expect(isValidZipBuffer(textFile)).toBe(false);
    });

    test('should reject PDF signature', () => {
      const pdf = Buffer.from([0x25, 0x50, 0x44, 0x46]); // %PDF
      expect(isValidZipBuffer(pdf)).toBe(false);
    });
  });

  describe('isTextContent', () => {
    test('should identify string as text', () => {
      expect(isTextContent('Hello World')).toBe(true);
    });

    test('should identify UTF-8 buffer as text', () => {
      const textBuffer = Buffer.from('Hello World', 'utf8');
      expect(isTextContent(textBuffer)).toBe(true);
    });

    test('should identify buffer with null bytes as binary', () => {
      const binaryBuffer = Buffer.from([0x48, 0x65, 0x00, 0x6c, 0x6c]); // He\0ll
      expect(isTextContent(binaryBuffer)).toBe(false);
    });

    test('should handle unicode text', () => {
      const unicodeText = 'Hello 世界 🌍';
      expect(isTextContent(unicodeText)).toBe(true);

      const unicodeBuffer = Buffer.from(unicodeText, 'utf8');
      expect(isTextContent(unicodeBuffer)).toBe(true);
    });

    test('should identify binary data as non-text', () => {
      const binaryData = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x00, 0x00]);
      expect(isTextContent(binaryData)).toBe(false);
    });

    test('should handle empty string', () => {
      expect(isTextContent('')).toBe(true);
    });

    test('should handle empty buffer', () => {
      expect(isTextContent(Buffer.from([]))).toBe(true);
    });
  });

  describe('sanitizeHyperlinkUrl', () => {
    test('should strip Chrome extension prefix (Adobe Acrobat)', () => {
      const url = 'chrome-extension://efaidnbmnnnibpcajpcglclefindmkaj/https://example.com/doc.pdf';
      const result = sanitizeHyperlinkUrl(url);
      expect(result).not.toBeNull();
      expect(result!.url).toBe('https://example.com/doc.pdf');
      expect(result!.fixes).toContain('Stripped Chrome extension URL prefix');
    });

    test('should strip generic Chrome extension prefix (random 32-char ID)', () => {
      const url = 'chrome-extension://abcdefghijklmnopqrstuvwxyzabcdef/https://example.com';
      const result = sanitizeHyperlinkUrl(url);
      expect(result).not.toBeNull();
      expect(result!.url).toBe('https://example.com');
    });

    test('should strip Firefox moz-extension prefix', () => {
      const url = 'moz-extension://a1b2c3d4-e5f6-7890-abcd-ef1234567890/https://example.com';
      const result = sanitizeHyperlinkUrl(url);
      expect(result).not.toBeNull();
      expect(result!.url).toBe('https://example.com');
      expect(result!.fixes).toContain('Stripped Firefox extension URL prefix');
    });

    test('should strip Edge extension prefix', () => {
      const url = 'extension://abcdefghijklmnopqrstuvwxyz123456/https://example.com';
      const result = sanitizeHyperlinkUrl(url);
      expect(result).not.toBeNull();
      expect(result!.url).toBe('https://example.com');
      expect(result!.fixes).toContain('Stripped Edge extension URL prefix');
    });

    test('should fix broken protocol after stripping (https:/ -> https://)', () => {
      const url = 'chrome-extension://efaidnbmnnnibpcajpcglclefindmkaj/https:/example.com/path';
      const result = sanitizeHyperlinkUrl(url);
      expect(result).not.toBeNull();
      expect(result!.url).toBe('https://example.com/path');
      expect(result!.fixes).toHaveLength(2);
      expect(result!.fixes).toContain('Stripped Chrome extension URL prefix');
      expect(result!.fixes).toContain('Fixed broken protocol (added missing slash)');
    });

    test('should return null for clean URLs (no changes needed)', () => {
      expect(sanitizeHyperlinkUrl('https://example.com')).toBeNull();
      expect(sanitizeHyperlinkUrl('http://example.com/path')).toBeNull();
      expect(sanitizeHyperlinkUrl('mailto:user@example.com')).toBeNull();
      expect(sanitizeHyperlinkUrl('')).toBeNull();
    });

    test('should handle combined prefix + broken protocol', () => {
      const url =
        'chrome-extension://efaidnbmnnnibpcajpcglclefindmkaj/https:/aetnao365.sharepoint.com/sites/test';
      const result = sanitizeHyperlinkUrl(url);
      expect(result).not.toBeNull();
      expect(result!.url).toBe('https://aetnao365.sharepoint.com/sites/test');
    });

    test('should fix broken http:/ protocol', () => {
      const url = 'chrome-extension://efaidnbmnnnibpcajpcglclefindmkaj/http:/example.com';
      const result = sanitizeHyperlinkUrl(url);
      expect(result).not.toBeNull();
      expect(result!.url).toBe('http://example.com');
    });
  });
});
