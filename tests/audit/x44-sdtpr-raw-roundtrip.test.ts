/**
 * CT_SdtPr (ECMA-376 §17.5.2.38) has many children the object model does
 * not cover (w:rPr placeholder formatting, w:temporary, w15:appearance,
 * w15:color, ...), and w:sdtEndPr is not modeled at all. toXML() used to
 * rebuild sdtPr from the modeled subset on every save, silently dropping
 * everything else on a plain load -> save. The parser now captures the
 * original sdtPr/sdtEndPr markup and re-emits it verbatim until a modeled
 * property is mutated, at which point the rebuild path takes over.
 */
import { Document } from '../../src/core/Document';
import { StructuredDocumentTag } from '../../src/elements/StructuredDocumentTag';
import { ZipHandler } from '../../src/zip/ZipHandler';

const NESTED_SDT =
  `<w:sdt>` +
  `<w:sdtPr><w:id w:val="555"/><w15:appearance w15:val="hidden"/></w:sdtPr>` +
  `<w:sdtContent><w:p><w:r><w:t>NESTED-BODY</w:t></w:r></w:p></w:sdtContent>` +
  `</w:sdt>`;

const BLOCK_SDT =
  `<w:sdt>` +
  `<w:sdtPr>` +
  `<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>` +
  `<w:id w:val="123456789"/>` +
  `<w:tag w:val="orig-tag"/>` +
  `<w:temporary/>` +
  `<w15:appearance w15:val="tags"/>` +
  `<w15:color w:val="00FF00"/>` +
  `<w:richText/>` +
  `</w:sdtPr>` +
  `<w:sdtEndPr><w:rPr><w:i/></w:rPr></w:sdtEndPr>` +
  `<w:sdtContent>` +
  `<w:p><w:r><w:t>SDT-BODY</w:t></w:r></w:p>` +
  NESTED_SDT +
  `</w:sdtContent>` +
  `</w:sdt>`;

async function buildDocxWithSdt(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('sdtPr raw round-trip');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);

  const docXml = zip.getFileAsString('word/document.xml')!;
  const updatedDoc = docXml.includes('<w:sectPr')
    ? docXml.replace('<w:sectPr', `${BLOCK_SDT}<w:sectPr`)
    : docXml.replace('</w:body>', `${BLOCK_SDT}</w:body>`);
  zip.updateFile('word/document.xml', updatedDoc);

  return zip.toBuffer();
}

async function saveAndReadDocXml(doc: Document): Promise<string> {
  const out = await doc.toBuffer();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(out);
  return zip.getFileAsString('word/document.xml')!;
}

describe('w:sdtPr / w:sdtEndPr raw round-trip', () => {
  it('preserves unmodeled sdtPr children on a plain load -> save', async () => {
    const doc = await Document.loadFromBuffer(await buildDocxWithSdt());
    try {
      const docXml = await saveAndReadDocXml(doc);

      expect(docXml).toContain('<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>');
      expect(docXml).toContain('<w:temporary/>');
      expect(docXml).toContain('<w15:appearance w15:val="tags"/>');
      expect(docXml).toContain('<w15:color w:val="00FF00"/>');
      expect(docXml).toContain('<w:tag w:val="orig-tag"/>');
    } finally {
      doc.dispose();
    }
  });

  it('preserves w:sdtEndPr on a plain load -> save', async () => {
    const doc = await Document.loadFromBuffer(await buildDocxWithSdt());
    try {
      const docXml = await saveAndReadDocXml(doc);

      expect(docXml).toContain('<w:sdtEndPr><w:rPr><w:i/></w:rPr></w:sdtEndPr>');
      // CT_SdtBlock order: sdtPr, sdtEndPr, sdtContent
      const endPrPos = docXml.indexOf('<w:sdtEndPr>');
      const contentPos = docXml.indexOf('<w:sdtContent>');
      expect(endPrPos).toBeGreaterThan(-1);
      expect(contentPos).toBeGreaterThan(endPrPos);
    } finally {
      doc.dispose();
    }
  });

  it('preserves unmodeled sdtPr children of nested SDTs', async () => {
    const doc = await Document.loadFromBuffer(await buildDocxWithSdt());
    try {
      const docXml = await saveAndReadDocXml(doc);

      expect(docXml).toContain('w15:val="hidden"');
      expect(docXml).toContain('NESTED-BODY');
    } finally {
      doc.dispose();
    }
  });

  it('rebuilds sdtPr from the model once a modeled property is mutated', async () => {
    const doc = await Document.loadFromBuffer(await buildDocxWithSdt());
    try {
      const sdt = doc
        .getBodyElements()
        .find((el) => el instanceof StructuredDocumentTag) as StructuredDocumentTag;
      expect(sdt).toBeDefined();
      expect(sdt.getTag()).toBe('orig-tag');

      sdt.setTag('mutated-tag');
      const docXml = await saveAndReadDocXml(doc);

      // Stale raw markup must not shadow the mutation
      expect(docXml).toContain('w:val="mutated-tag"');
      expect(docXml).not.toContain('orig-tag');
    } finally {
      doc.dispose();
    }
  });

  it('survives a second round-trip (re-emitted markup parses again)', async () => {
    const doc1 = await Document.loadFromBuffer(await buildDocxWithSdt());
    let buf2: Buffer;
    try {
      buf2 = await doc1.toBuffer();
    } finally {
      doc1.dispose();
    }

    const doc2 = await Document.loadFromBuffer(buf2);
    try {
      const docXml = await saveAndReadDocXml(doc2);
      expect(docXml).toContain('<w15:appearance w15:val="tags"/>');
      expect(docXml).toContain('<w:sdtEndPr><w:rPr><w:i/></w:rPr></w:sdtEndPr>');

      const sdt = doc2
        .getBodyElements()
        .find((el) => el instanceof StructuredDocumentTag) as StructuredDocumentTag;
      expect(sdt.getTag()).toBe('orig-tag');
    } finally {
      doc2.dispose();
    }
  });
});
