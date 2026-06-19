/**
 * mergeAppPropsWithOriginal() must insert Manager/Application/AppVersion/Company
 * when the tag is absent from the original docProps/app.xml.
 *
 * Word-authored app.xml omits these optional extended-properties unless they
 * were previously set. The merge path is replace-only: if the tag is missing it
 * was silently skipped, so setManager()/setApplication()/setAppVersion() on a
 * loaded document was a no-op in the saved file even though the in-memory getter
 * still reported the value. The fix adds an insertion fallback before
 * </Properties> for each absent tag.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLParser } from '../../src/xml/XMLParser';

/**
 * Builds a minimal .docx whose app.xml contains only Application/AppVersion and
 * NO Manager/Company tag — the typical Word-authored shape.
 */
async function createDocxWithAppXml(appXml: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();

  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`
  );

  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`
  );

  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>
</w:document>`
  );

  zipHandler.addFile('docProps/app.xml', appXml);

  return await zipHandler.toBuffer();
}

const APP_XML_NO_MANAGER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office Word</Application>
  <AppVersion>16.0000</AppVersion>
</Properties>`;

async function saveAndReadAppXml(doc: Document): Promise<string> {
  const outBuffer = await doc.toBuffer();
  const outZip = new ZipHandler();
  await outZip.loadFromBuffer(outBuffer);
  const xml = outZip.getFileAsString('docProps/app.xml');
  expect(xml).toBeDefined();
  return xml as string;
}

describe('mergeAppPropsWithOriginal inserts absent extended-property tags', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('setManager() is persisted when <Manager> is absent from the original app.xml', async () => {
    const buffer = await createDocxWithAppXml(APP_XML_NO_MANAGER);
    doc = await Document.loadFromBuffer(buffer);

    doc.setManager('Jane Boss');

    const xml = await saveAndReadAppXml(doc);
    expect(xml).toContain('<Manager>Jane Boss</Manager>');
  });

  it('setApplication() is persisted when <Application> is absent', async () => {
    const appXmlNoApplication = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <AppVersion>16.0000</AppVersion>
</Properties>`;
    const buffer = await createDocxWithAppXml(appXmlNoApplication);
    doc = await Document.loadFromBuffer(buffer);

    doc.setApplication('docxmlater');

    const xml = await saveAndReadAppXml(doc);
    expect(xml).toContain('<Application>docxmlater</Application>');
  });

  it('setAppVersion() is persisted when <AppVersion> is absent', async () => {
    const appXmlNoAppVersion = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office Word</Application>
</Properties>`;
    const buffer = await createDocxWithAppXml(appXmlNoAppVersion);
    doc = await Document.loadFromBuffer(buffer);

    doc.setAppVersion('2.5.0');

    const xml = await saveAndReadAppXml(doc);
    expect(xml).toContain('<AppVersion>2.5.0</AppVersion>');
  });

  it('setCompany() is persisted when <Company> is fully absent (not just self-closing)', async () => {
    const buffer = await createDocxWithAppXml(APP_XML_NO_MANAGER);
    doc = await Document.loadFromBuffer(buffer);

    doc.setCompany('Acme Corp');

    const xml = await saveAndReadAppXml(doc);
    expect(xml).toContain('<Company>Acme Corp</Company>');
  });

  it('the original preserved properties survive the insertion', async () => {
    const buffer = await createDocxWithAppXml(APP_XML_NO_MANAGER);
    doc = await Document.loadFromBuffer(buffer);

    doc.setManager('Jane Boss');

    const xml = await saveAndReadAppXml(doc);
    // Inserted value plus the untouched originals
    expect(xml).toContain('<Manager>Jane Boss</Manager>');
    expect(xml).toContain('<Application>Microsoft Office Word</Application>');
    expect(xml).toContain('<AppVersion>16.0000</AppVersion>');
  });

  it('inserts the tag inside <Properties> and keeps the part well-formed', async () => {
    const buffer = await createDocxWithAppXml(APP_XML_NO_MANAGER);
    doc = await Document.loadFromBuffer(buffer);

    doc.setManager('Jane Boss');

    const xml = await saveAndReadAppXml(doc);
    // The inserted element must live INSIDE the root, not after </Properties>.
    const managerAt = xml.indexOf('<Manager>');
    const closeAt = xml.indexOf('</Properties>');
    expect(managerAt).toBeGreaterThan(-1);
    expect(closeAt).toBeGreaterThan(-1);
    expect(managerAt).toBeLessThan(closeAt);

    // Exactly one Properties root and a single inserted Manager (no duplicates),
    // and the part re-parses to an object exposing the inserted value.
    expect(xml.match(/<\/Properties>/g)?.length).toBe(1);
    expect(xml.match(/<Manager>/g)?.length).toBe(1);
    const parsed = XMLParser.parseToObject(xml) as any;
    expect(parsed.Properties).toBeDefined();
    expect(parsed.Properties.Manager).toBe('Jane Boss');
  });
});
