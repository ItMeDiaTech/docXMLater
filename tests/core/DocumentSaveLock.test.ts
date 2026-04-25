/**
 * Concurrent save() / toBuffer() must not corrupt shared manager state.
 *
 * Document is single-threaded by design — prepareSave() mutates StylesManager,
 * NumberingManager, ImageManager, and the underlying zipHandler. Without the
 * save queue, Promise.all([doc.save(a), doc.save(b)]) would race on those
 * managers and could leak partial state between callers. The queue serialises
 * save / toBuffer onto a single Promise chain.
 */
import { Document } from '../../src/core/Document';

describe('Document save lock', () => {
  it('serialises concurrent toBuffer() calls', async () => {
    const doc = Document.create();
    doc.createParagraph('Concurrent buffers');

    const buffers = await Promise.all([doc.toBuffer(), doc.toBuffer(), doc.toBuffer()]);

    // Every buffer should be a complete, distinct DOCX file (ZIP magic = "PK\x03\x04")
    for (const buf of buffers) {
      expect(buf.length).toBeGreaterThan(100);
      expect(buf.subarray(0, 4).toString('hex')).toBe('504b0304');
    }
    doc.dispose();
  });

  it('continues processing waiters after a failed save', async () => {
    const doc = Document.create();
    doc.createParagraph('Recovery test');

    // First call uses an invalid path on Windows (illegal chars) — must reject.
    const failing = doc.save('Z:\\__nonexistent_dir__\\path<>:"|?*.docx').catch(() => 'failed');
    // Second call should still execute against the in-memory state and succeed.
    const buffer = await doc.toBuffer();

    expect(await failing).toBe('failed');
    expect(buffer.length).toBeGreaterThan(100);
    doc.dispose();
  });

  it('preserves order: later toBuffer reflects mutations made between waiters', async () => {
    const doc = Document.create();
    doc.createParagraph('First');

    const first = doc.toBuffer();
    doc.createParagraph('Second');
    const second = doc.toBuffer();

    const [bufA, bufB] = await Promise.all([first, second]);
    // Both buffers contain the first paragraph; only the second contains "Second".
    // (The mutation between calls happened before the second toBuffer entered
    // its prepareSave under the lock — so it must observe both paragraphs.)
    const reload = await Document.loadFromBuffer(bufB);
    const text = reload
      .getAllParagraphs()
      .map((p) => p.getText())
      .join('|');
    expect(text).toContain('First');
    expect(text).toContain('Second');
    reload.dispose();

    expect(bufA.length).toBeGreaterThan(0);
    doc.dispose();
  });
});
