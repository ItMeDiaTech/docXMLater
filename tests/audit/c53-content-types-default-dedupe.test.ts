/**
 * Regression: [Content_Types].xml merge must emit at most one <Default> per
 * extension and at most one <Override> per part name (ECMA-376 Part 2
 * §10.1.2.2). A loaded document that declared the same extension with a
 * different-but-valid MIME string (e.g. image/emf vs the framework's
 * image/x-emf) previously produced duplicate <Default> elements, which
 * strict OPC consumers reject. The original document's declared content
 * type must win for round-trip fidelity.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function buildDocxWithEmfVariantMime(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('Original');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);

  // Unreferenced media file so generation's word/media scan registers
  // the framework's hardcoded 'emf|image/x-emf' default.
  zip.addFile('word/media/image1.emf', Buffer.from([0x01, 0x00, 0x00, 0x00]), {
    binary: true,
  });

  // Original document declares the same extension with the IANA-registered
  // MIME some producers use.
  const ct = zip.getFileAsString('[Content_Types].xml')!;
  const updated = ct.replace(
    '</Types>',
    '<Default Extension="emf" ContentType="image/emf"/></Types>'
  );
  zip.updateFile('[Content_Types].xml', updated);

  return zip.toBuffer();
}

describe('[Content_Types].xml merge dedupes by extension and part name', () => {
  it('emits a single <Default> per extension, keeping the original MIME type', async () => {
    const buffer1 = await buildDocxWithEmfVariantMime();
    const doc = await Document.loadFromBuffer(buffer1);
    const buffer2 = await doc.toBuffer();
    doc.dispose();

    const out = new ZipHandler();
    await out.loadFromBuffer(buffer2);
    const ct = out.getFileAsString('[Content_Types].xml')!;

    const emfDefaults = ct.match(/<Default Extension="emf"/g) || [];
    expect(emfDefaults).toHaveLength(1);
    expect(ct).toContain('<Default Extension="emf" ContentType="image/emf"/>');
    expect(ct).not.toContain('image/x-emf');
  });

  it('emits a single <Override> per part name, keeping the original content type', async () => {
    const seed = Document.create();
    seed.createParagraph('Original');
    const base = await seed.toBuffer();
    seed.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(base);

    // Original declares a macro-enabled main part; the generated override
    // for /word/document.xml uses the standard content type.
    const ct = zip.getFileAsString('[Content_Types].xml')!;
    const updated = ct.replace(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
      'application/vnd.ms-word.document.macroEnabled.main+xml'
    );
    zip.updateFile('[Content_Types].xml', updated);
    const buffer1 = await zip.toBuffer();

    const doc = await Document.loadFromBuffer(buffer1);
    const buffer2 = await doc.toBuffer();
    doc.dispose();

    const out = new ZipHandler();
    await out.loadFromBuffer(buffer2);
    const outCt = out.getFileAsString('[Content_Types].xml')!;

    const docOverrides = outCt.match(/<Override PartName="\/word\/document\.xml"/g) || [];
    expect(docOverrides).toHaveLength(1);
    expect(outCt).toContain('application/vnd.ms-word.document.macroEnabled.main+xml');
  });

  it('does not duplicate defaults on a plain round-trip', async () => {
    const seed = Document.create();
    seed.createParagraph('Plain');
    const buffer1 = await seed.toBuffer();
    seed.dispose();

    const doc = await Document.loadFromBuffer(buffer1);
    const buffer2 = await doc.toBuffer();
    doc.dispose();

    const out = new ZipHandler();
    await out.loadFromBuffer(buffer2);
    const ct = out.getFileAsString('[Content_Types].xml')!;

    const extensions = [...ct.matchAll(/<Default Extension="([^"]+)"/g)].map((m) =>
      m[1]!.toLowerCase()
    );
    expect(new Set(extensions).size).toBe(extensions.length);
  });
});
