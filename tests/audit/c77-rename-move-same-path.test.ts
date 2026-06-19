/**
 * Regression: ZipHandler.renameFile / moveFile must not destroy the file when
 * source and destination normalize to the same archive key.
 *
 * Pre-fix, addFile(newPath) overwrote the single Map entry and the subsequent
 * removeFile(oldPath) deleted it, so the file vanished while the method still
 * returned true.
 */

import { ZipHandler } from '../../src/zip/ZipHandler';

describe('C77: ZipHandler same-path rename/move guard', () => {
  let handler: ZipHandler;

  beforeEach(() => {
    handler = new ZipHandler();
  });

  test('renameFile with identical string paths is a no-op that keeps the file', () => {
    handler.addFile('word/document.xml', '<w:document/>');

    const renamed = handler.renameFile('word/document.xml', 'word/document.xml');

    expect(renamed).toBe(true);
    expect(handler.hasFile('word/document.xml')).toBe(true);
    expect(handler.getFileAsString('word/document.xml')).toBe('<w:document/>');
  });

  test('renameFile with backslash variant of the same path keeps the file', () => {
    const buffer = Buffer.from([0x89, 0x50, 0x4e, 0x47]);
    handler.addFile('word/media/image1.png', buffer, { binary: true });

    const renamed = handler.renameFile('word/media/image1.png', 'word\\media\\image1.png');

    expect(renamed).toBe(true);
    expect(handler.hasFile('word/media/image1.png')).toBe(true);
    expect(handler.getFileAsBuffer('word/media/image1.png')).toEqual(buffer);
  });

  test('moveFile with backslash variant of the same path keeps the file', () => {
    const buffer = Buffer.from([0x89, 0x50, 0x4e, 0x47]);
    handler.addFile('word/media/image1.png', buffer, { binary: true });

    const moved = handler.moveFile('word/media/image1.png', 'word\\media\\image1.png');

    expect(moved).toBe(true);
    expect(handler.hasFile('word/media/image1.png')).toBe(true);
    expect(handler.getFileAsBuffer('word/media/image1.png')).toEqual(buffer);
  });

  test('distinct-path rename still moves the file (no false positive on the guard)', () => {
    handler.addFile('old.txt', 'Content');

    const renamed = handler.renameFile('old.txt', 'new.txt');

    expect(renamed).toBe(true);
    expect(handler.hasFile('new.txt')).toBe(true);
    expect(handler.hasFile('old.txt')).toBe(false);
  });
});
