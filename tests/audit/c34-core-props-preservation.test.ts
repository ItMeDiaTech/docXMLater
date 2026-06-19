/**
 * docProps/core.xml round-trip preservation.
 *
 * The parser extracts only a fixed subset of OPC core properties
 * (title/subject/creator/keywords/description/category/contentStatus/language/
 * revision/created/modified). Other valid optional elements — cp:lastPrinted,
 * dc:identifier, cp:version — are never parsed. Previously updateCoreProps()
 * unconditionally regenerated core.xml from that subset on every save, silently
 * dropping the unparsed elements. The fix stores _originalCorePropsXml at load
 * and a _corePropsModified dirty flag so the original is preserved verbatim when
 * unchanged, and only the changed fields are merged in otherwise.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

const CORE_XML_WITH_EXTRA = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Original Title</dc:title>
  <dc:creator>Original Author</dc:creator>
  <cp:lastModifiedBy>Original Author</cp:lastModifiedBy>
  <cp:revision>4</cp:revision>
  <cp:lastPrinted>2024-01-15T09:30:00Z</cp:lastPrinted>
  <dc:identifier>doc-uuid-12345</dc:identifier>
  <cp:version>2.1</cp:version>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-10T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;

async function createDocxWithCoreXml(coreXml: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();

  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`
  );

  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`
  );

  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>
</w:document>`
  );

  zipHandler.addFile('docProps/core.xml', coreXml);

  return await zipHandler.toBuffer();
}

async function saveAndReadCoreXml(doc: Document): Promise<string> {
  const outBuffer = await doc.toBuffer();
  const outZip = new ZipHandler();
  await outZip.loadFromBuffer(outBuffer);
  const xml = outZip.getFileAsString('docProps/core.xml');
  expect(xml).toBeDefined();
  return xml as string;
}

describe('docProps/core.xml preservation (round-trip fidelity)', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('preserves unparsed core properties on a plain load→save round-trip', async () => {
    const buffer = await createDocxWithCoreXml(CORE_XML_WITH_EXTRA);
    doc = await Document.loadFromBuffer(buffer);

    // No property setter called — original must survive verbatim.
    const xml = await saveAndReadCoreXml(doc);
    expect(xml).toContain('<cp:lastPrinted>2024-01-15T09:30:00Z</cp:lastPrinted>');
    expect(xml).toContain('<dc:identifier>doc-uuid-12345</dc:identifier>');
    expect(xml).toContain('<cp:version>2.1</cp:version>');
  });

  it('keeps unparsed core properties when a parsed field is changed (merge path)', async () => {
    const buffer = await createDocxWithCoreXml(CORE_XML_WITH_EXTRA);
    doc = await Document.loadFromBuffer(buffer);

    doc.setTitle('Updated Title');

    const xml = await saveAndReadCoreXml(doc);
    // Changed field is applied...
    expect(xml).toContain('<dc:title>Updated Title</dc:title>');
    // ...and the unparsed elements still survive.
    expect(xml).toContain('<cp:lastPrinted>2024-01-15T09:30:00Z</cp:lastPrinted>');
    expect(xml).toContain('<dc:identifier>doc-uuid-12345</dc:identifier>');
    expect(xml).toContain('<cp:version>2.1</cp:version>');
    // Untouched parsed fields preserved too.
    expect(xml).toContain('<dc:creator>Original Author</dc:creator>');
  });

  it('inserts a parsed field that was absent from the original core.xml', async () => {
    const coreXmlNoSubject = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>T</dc:title>
  <cp:lastPrinted>2024-01-15T09:30:00Z</cp:lastPrinted>
</cp:coreProperties>`;
    const buffer = await createDocxWithCoreXml(coreXmlNoSubject);
    doc = await Document.loadFromBuffer(buffer);

    doc.setSubject('Brand New Subject');

    const xml = await saveAndReadCoreXml(doc);
    expect(xml).toContain('<dc:subject>Brand New Subject</dc:subject>');
    expect(xml).toContain('<cp:lastPrinted>2024-01-15T09:30:00Z</cp:lastPrinted>');
  });
});
