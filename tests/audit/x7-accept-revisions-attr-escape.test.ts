/**
 * The raw-XML revision acceptor (DOM path) re-serializes every part that
 * contains revision markup. Its objectToXml escapes attribute values on
 * output, so the parser must hand it decoded values: when attribute values
 * were stored still-escaped, w:instr="HYPERLINK &quot;...&amp;...&quot;"
 * came back as &amp;quot;/&amp;amp; — Word then shows the literal entity
 * text, breaking field instructions, tooltips, and anchors in any part
 * that carries even a single tracked change.
 */
import { ZipHandler } from '../../src/zip/ZipHandler';
import { acceptAllRevisions } from '../../src/processors/acceptRevisions';

function makeZip(documentXml: string): ZipHandler {
  const zip = new ZipHandler();
  zip.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  zip.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zip.addFile('word/document.xml', documentXml);
  return zip;
}

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:ins w:id="1" w:author="Smith &amp; Jones" w:date="2026-06-01T10:00:00Z">
        <w:r><w:t>inserted</w:t></w:r>
      </w:ins>
      <w:fldSimple w:instr=" HYPERLINK &quot;https://example.com/?a=1&amp;b=2&quot; ">
        <w:r><w:t>link</w:t></w:r>
      </w:fldSimple>
      <w:hyperlink r:id="rId4" w:tooltip="Q&amp;A &quot;answers&quot;">
        <w:r><w:t>anchor</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;

describe('acceptAllRevisions (raw-XML DOM path) — entity-bearing attributes', () => {
  it('accepts the insertion and keeps entity attributes single-escaped', async () => {
    const zip = makeZip(DOCUMENT_XML);
    await acceptAllRevisions(zip);
    const after = zip.getFileAsString('word/document.xml')!;

    // Revision accepted: wrapper gone, content kept.
    expect(after).not.toMatch(/<w:ins\b/);
    expect(after).toContain('inserted');

    // Attributes elsewhere in the part survive exactly single-escaped.
    expect(after).toContain('w:instr=" HYPERLINK &quot;https://example.com/?a=1&amp;b=2&quot; "');
    expect(after).toContain('w:tooltip="Q&amp;A &quot;answers&quot;"');
    expect(after).not.toContain('&amp;quot;');
    expect(after).not.toContain('&amp;amp;');
  });

  it('is stable across repeated accept passes (no compounding escapes)', async () => {
    const zip = makeZip(DOCUMENT_XML);
    await acceptAllRevisions(zip);
    const first = zip.getFileAsString('word/document.xml')!;

    // Re-introduce a tracked change so the part is re-serialized again.
    const reinjected = first.replace(
      '<w:p>',
      '<w:p><w:ins w:id="2" w:author="B" w:date="2026-06-01T11:00:00Z"><w:r><w:t>again</w:t></w:r></w:ins>'
    );
    const zip2 = makeZip(reinjected);
    await acceptAllRevisions(zip2);
    const second = zip2.getFileAsString('word/document.xml')!;

    expect(second).toContain('w:instr=" HYPERLINK &quot;https://example.com/?a=1&amp;b=2&quot; "');
    expect(second).not.toContain('&amp;quot;');
    expect(second).not.toContain('&amp;amp;');
  });
});
