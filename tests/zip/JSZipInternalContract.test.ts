/**
 * Canary contract test for JSZip's undocumented per-entry `_data` field.
 *
 * ZipReader uses `_data.compressedSize` / `_data.uncompressedSize` as an early-reject
 * hint for its resource-limit guards (see {@link JSZipObjectPrivate} in ZipReader.ts).
 * The field is internal to JSZip and only guaranteed by manual verification against the
 * pinned range (`jszip ^3.10.1`, verified 3.10.x). A jszip upgrade that renames or
 * removes it would silently turn the early-reject/ratio hints into no-ops — the measured
 * decompressed-byte accounting still enforces the budgets, but the cheap pre-check would
 * be lost without any signal. [R-F6]
 *
 * This test reproduces exactly how ZipReader observes the field — load a normally-built
 * archive with `JSZip.loadAsync` and inspect `zip.files[path]._data` — and fails loudly
 * if the shape changes, so the degradation surfaces at upgrade time rather than in prod.
 */

import JSZip from 'jszip';

/** Mirrors the (non-exported) internal shape ZipReader casts each entry to. */
interface JSZipObjectPrivate {
  _data?: {
    compressedSize?: number;
    uncompressedSize?: number;
  };
}

describe('JSZip internal `_data` contract (canary)', () => {
  test('a loaded, DEFLATE-built entry exposes a numeric _data.compressedSize and uncompressedSize', async () => {
    // Build a sizable, compressible entry so the archive actually stores a deflate stream.
    const source = new JSZip();
    source.file('payload.txt', 'x'.repeat(64 * 1024));
    const buffer = await source.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });

    // Observe the field the same way ZipReader does: via loadAsync, not the freshly-added
    // entry (compressed sizes only exist on the CompressedObject produced by a load).
    const zip = await JSZip.loadAsync(buffer);
    const entry = zip.files['payload.txt'];
    expect(entry).toBeDefined();

    const internalData = (entry as unknown as JSZipObjectPrivate)._data;
    expect(internalData).toBeDefined();
    expect(typeof internalData!.compressedSize).toBe('number');
    expect(internalData!.compressedSize).toBeGreaterThan(0);
    expect(typeof internalData!.uncompressedSize).toBe('number');
    expect(internalData!.uncompressedSize).toBe(64 * 1024);
  });

  test('a STORED (uncompressed) entry still reports both sizes', async () => {
    const source = new JSZip();
    source.file('stored.bin', Buffer.alloc(2048, 0x41));
    const buffer = await source.generateAsync({ type: 'nodebuffer', compression: 'STORE' });

    const zip = await JSZip.loadAsync(buffer);
    const internalData = (zip.files['stored.bin'] as unknown as JSZipObjectPrivate)._data;
    expect(internalData).toBeDefined();
    expect(typeof internalData!.compressedSize).toBe('number');
    expect(typeof internalData!.uncompressedSize).toBe('number');
    expect(internalData!.uncompressedSize).toBe(2048);
  });
});
