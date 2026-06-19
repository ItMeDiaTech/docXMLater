/**
 * Binary parts whose extensions are not images/fonts — embedded OLE packages
 * (word/embeddings/*.xlsx), obfuscated embedded fonts (word/fonts/*.odttf),
 * and compressed metafiles (word/media/*.emz) — must survive load → save
 * byte-for-byte. Extracting them as UTF-8 strings is a lossy decode: invalid
 * sequences become U+FFFD (EF BF BD), silently destroying the part.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

// Deliberately invalid UTF-8 (0x80, 0xff, 0xfe, lone 0xc3) — a lossy string
// decode replaces these bytes with U+FFFD, so byte equality detects corruption.
const BINARY_PAYLOAD = Buffer.from([
  0x50, 0x4b, 0x03, 0x04, 0x80, 0xff, 0xfe, 0xc3, 0x28, 0xb5, 0x2f, 0xfd, 0x00, 0x9d, 0xe5,
]);

const BINARY_PARTS = [
  'word/embeddings/Microsoft_Excel_Worksheet1.xlsx',
  'word/fonts/font1.odttf',
  'word/media/image1.emz',
];

async function buildDocxWithBinaryParts(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('Binary part round-trip');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);

  for (const path of BINARY_PARTS) {
    zip.addFile(path, BINARY_PAYLOAD, { binary: true });
  }

  const ct = zip.getFileAsString('[Content_Types].xml')!;
  const updatedCt = ct.replace(
    '</Types>',
    `<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>` +
      `<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>` +
      `<Default Extension="emz" ContentType="image/x-emz"/>` +
      `</Types>`
  );
  zip.updateFile('[Content_Types].xml', updatedCt);

  return zip.toBuffer();
}

describe('Binary part round-trip for unlisted extensions', () => {
  it('preserves xlsx/odttf/emz bytes through ZipHandler load + toBuffer', async () => {
    const buf1 = await buildDocxWithBinaryParts();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(buf1);
    const buf2 = await zip.toBuffer();

    const out = new ZipHandler();
    await out.loadFromBuffer(buf2);
    for (const path of BINARY_PARTS) {
      const content = out.getFileAsBuffer(path);
      expect(content).toBeDefined();
      expect(content!.equals(BINARY_PAYLOAD)).toBe(true);
    }
  });

  it('preserves xlsx/odttf/emz bytes through Document load + toBuffer', async () => {
    const buf1 = await buildDocxWithBinaryParts();

    const doc = await Document.loadFromBuffer(buf1);
    try {
      const buf2 = await doc.toBuffer();

      const out = new ZipHandler();
      await out.loadFromBuffer(buf2);
      for (const path of BINARY_PARTS) {
        const content = out.getFileAsBuffer(path);
        expect(content).toBeDefined();
        expect(content!.equals(BINARY_PAYLOAD)).toBe(true);
      }
    } finally {
      doc.dispose();
    }
  });

  it('preserves bytes for extensionless parts', async () => {
    const seed = Document.create();
    seed.createParagraph('Extensionless part');
    const base = await seed.toBuffer();
    seed.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(base);
    zip.addFile('word/embeddings/payload', BINARY_PAYLOAD, { binary: true });
    const buf1 = await zip.toBuffer();

    const reloaded = new ZipHandler();
    await reloaded.loadFromBuffer(buf1);
    const buf2 = await reloaded.toBuffer();

    const out = new ZipHandler();
    await out.loadFromBuffer(buf2);
    const content = out.getFileAsBuffer('word/embeddings/payload');
    expect(content).toBeDefined();
    expect(content!.equals(BINARY_PAYLOAD)).toBe(true);
  });
});
