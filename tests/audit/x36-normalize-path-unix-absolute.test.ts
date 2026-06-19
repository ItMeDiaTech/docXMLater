/**
 * Tests for normalizePath's Unix absolute-path contract
 *
 * Unix absolute paths are sanitized to archive-relative paths by stripping
 * leading slashes (the path becomes a harmless ZIP entry key), while Windows
 * drive-letter paths are rejected. This asymmetry is the documented contract;
 * normalizePath must never throw for a leading-slash path.
 *
 * normalizePath previously carried a dead Unix absolute-path guard
 * (`if (path.startsWith('/') && normalized.startsWith('/')) throw ...`) that
 * could never fire because leading slashes are stripped before the check, plus
 * a JSDoc/@throws contract that wrongly claimed `/etc/`-style paths are
 * rejected. The behavioral assertions below pin the sanitize-not-reject
 * contract; the source assertions pin removal of the unreachable branch so the
 * code no longer contradicts its own documentation.
 */

import { normalizePath } from '../../src/utils/validation';

describe('normalizePath Unix absolute paths', () => {
  test('sanitizes Unix absolute paths to relative instead of rejecting them', () => {
    expect(normalizePath('/etc/passwd')).toBe('etc/passwd');
    expect(normalizePath('//etc/passwd')).toBe('etc/passwd');
    expect(normalizePath('///word/document.xml')).toBe('word/document.xml');
  });

  test('sanitizes backslash-rooted paths the same way', () => {
    expect(normalizePath('\\word\\document.xml')).toBe('word/document.xml');
  });

  test('still rejects absolute Windows drive-letter paths', () => {
    expect(() => normalizePath('C:/Windows/System32')).toThrow('absolute Windows path');
    expect(() => normalizePath('C:\\Windows\\System32')).toThrow('absolute Windows path');
  });

  test('still rejects path traversal regardless of leading slash', () => {
    expect(() => normalizePath('/../etc/passwd')).toThrow('path traversal');
    expect(() => normalizePath('/word/../../etc/passwd')).toThrow('path traversal');
  });

  test('no longer carries the unreachable Unix absolute-path guard', () => {
    // The removed branch was `if (path.startsWith('/') && normalized.startsWith('/'))`
    // throwing an "absolute Unix path" error. It was unreachable because leading
    // slashes are stripped first, so it can only ever exist as dead code that
    // contradicts the function's own (corrected) documentation. Inspect the
    // compiled function body to ensure that dead branch and its throw are gone.
    const body = normalizePath.toString();
    expect(body).not.toContain('absolute Unix path');
    expect(body).not.toContain("normalized.startsWith('/')");
    expect(body).not.toContain('normalized.startsWith("/")');
  });
});
