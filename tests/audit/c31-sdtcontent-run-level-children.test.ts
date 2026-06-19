/**
 * w:sdtContent (CT_SdtContentBlock, ECMA-376 §17.5.2.42) admits w:customXml
 * plus EG_RunLevelElts — bookmark/comment/permission range markers,
 * w:proofErr, tracked-change wrappers, math — alongside w:p/w:tbl/w:sdt.
 * The parser used to reconstruct only w:p/w:tbl/w:sdt children, so every
 * other sdtContent child vanished when document.xml was regenerated on
 * save, unbalancing range pairs and dropping custom XML content.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

const BLOCK_SDT =
  `<w:sdt>` +
  `<w:sdtPr><w:id w:val="987654"/></w:sdtPr>` +
  `<w:sdtContent>` +
  `<w:bookmarkStart w:id="77" w:name="SdtScopedBookmark"/>` +
  `<w:p><w:r><w:t>SDT-PARA</w:t></w:r></w:p>` +
  `<w:bookmarkEnd w:id="77"/>` +
  `<w:customXml w:element="payload"><w:p><w:r><w:t>CX-TEXT</w:t></w:r></w:p></w:customXml>` +
  `</w:sdtContent>` +
  `</w:sdt>`;

async function buildDocxWithBlockSdt(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('SDT content-level children round-trip');
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

async function roundTrip(buf: Buffer): Promise<string> {
  const doc = await Document.loadFromBuffer(buf);
  try {
    const out = await doc.toBuffer();
    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    return zip.getFileAsString('word/document.xml')!;
  } finally {
    doc.dispose();
  }
}

describe('Block w:sdt content-level run-level children round-trip', () => {
  it('preserves bookmark range markers that are direct children of w:sdtContent', async () => {
    const docXml = await roundTrip(await buildDocxWithBlockSdt());

    expect(docXml).toContain('<w:bookmarkStart w:id="77" w:name="SdtScopedBookmark"/>');
    expect(docXml).toContain('<w:bookmarkEnd w:id="77"/>');
  });

  it('preserves w:customXml blocks that are direct children of w:sdtContent', async () => {
    const docXml = await roundTrip(await buildDocxWithBlockSdt());

    expect(docXml).toContain('<w:customXml w:element="payload">');
    // Text-only w:t must keep its wrapper through raw reconstruction
    expect(docXml).toContain('<w:t>CX-TEXT</w:t>');
  });

  it('keeps preserved children in their original position relative to paragraphs', async () => {
    const docXml = await roundTrip(await buildDocxWithBlockSdt());

    const contentPos = docXml.indexOf('<w:sdtContent>');
    const startPos = docXml.indexOf('<w:bookmarkStart w:id="77"');
    const paraPos = docXml.indexOf('SDT-PARA');
    const endPos = docXml.indexOf('<w:bookmarkEnd w:id="77"/>');
    const customXmlPos = docXml.indexOf('<w:customXml w:element="payload">');

    expect(contentPos).toBeGreaterThan(-1);
    expect(startPos).toBeGreaterThan(contentPos);
    expect(paraPos).toBeGreaterThan(startPos);
    expect(endPos).toBeGreaterThan(paraPos);
    expect(customXmlPos).toBeGreaterThan(endPos);
  });

  it('still parses the modeled paragraph content of the SDT', async () => {
    const doc = await Document.loadFromBuffer(await buildDocxWithBlockSdt());
    try {
      const text = doc
        .getAllParagraphs()
        .map((p) => p.getText())
        .join(' ');
      expect(text).toContain('SDT-PARA');
    } finally {
      doc.dispose();
    }
  });
});
