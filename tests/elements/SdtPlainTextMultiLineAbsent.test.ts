/**
 * `<w:sdtPr><w:text/>` — absent `w:multiLine` attribute must round-trip
 * as absent (not silently promoted to the spec-default value).
 *
 * Per ECMA-376 Part 1 §17.5.2.33 CT_SdtText, `w:multiLine` is an
 * OPTIONAL ST_OnOff attribute with a spec default of `false`. When the
 * source document omits the attribute, the in-memory model should
 * preserve "attribute absent" so the emitter doesn't write a gratuitous
 * `w:multiLine="0"` that the author never set.
 *
 * Bug: the parser at `DocumentParser.ts:7774` unconditionally invoked
 * `parseOnOffAttribute(textElement?.['@_w:multiLine'])` — when the attr
 * was undefined, that helper returned `false` (the default-value
 * parameter), so the parser always stored `multiLine: false`. The
 * emitter's `!== undefined` guard then wrote `w:multiLine="0"` on every
 * round-trip, adding spec-noise the author never wrote.
 *
 * Iteration 120 gates the parse on attribute presence so absent stays
 * absent, and the existing ST_OnOff literal coverage (for explicit-set
 * values) is preserved.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadAndResaveDocXml(xml: string): Promise<string> {
  const zipHandler = new ZipHandler();
  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zipHandler.addFile('word/document.xml', xml);
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const out = await doc.toBuffer();
  doc.dispose();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(out);
  const content = zip.getFile('word/document.xml')?.content;
  return content instanceof Buffer ? content.toString('utf8') : String(content);
}

function buildTextSdtDoc(textAttrFragment: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:sdt>
      <w:sdtPr>
        <w:id w:val="100"/>
        <w:text${textAttrFragment}/>
      </w:sdtPr>
      <w:sdtContent>
        <w:p><w:r><w:t>x</w:t></w:r></w:p>
      </w:sdtContent>
    </w:sdt>
    <w:p><w:r><w:t>d</w:t></w:r></w:p>
  </w:body>
</w:document>`;
}

function extractSdtText(xml: string): string {
  return xml.match(/<w:text[^/]*\/>/)?.[0] ?? '';
}

describe('<w:sdtPr><w:text> w:multiLine absence preservation', () => {
  it('omits w:multiLine on output when input had no w:multiLine attribute', async () => {
    const out = await loadAndResaveDocXml(buildTextSdtDoc(''));
    const textTag = extractSdtText(out);
    // Previously: output had `<w:text w:multiLine="0"/>` — the parser
    // promoted absent to `false`, the emitter wrote it as "0".
    expect(textTag).not.toMatch(/w:multiLine/);
  });

  it('preserves w:multiLine="1" explicit true', async () => {
    const out = await loadAndResaveDocXml(buildTextSdtDoc(' w:multiLine="1"'));
    expect(extractSdtText(out)).toMatch(/w:multiLine="1"/);
  });

  it('preserves w:multiLine="0" explicit false', async () => {
    const out = await loadAndResaveDocXml(buildTextSdtDoc(' w:multiLine="0"'));
    expect(extractSdtText(out)).toMatch(/w:multiLine="0"/);
  });

  it('normalises w:multiLine="on" to "1"', async () => {
    const out = await loadAndResaveDocXml(buildTextSdtDoc(' w:multiLine="on"'));
    expect(extractSdtText(out)).toMatch(/w:multiLine="1"/);
  });

  it('normalises w:multiLine="off" to "0"', async () => {
    const out = await loadAndResaveDocXml(buildTextSdtDoc(' w:multiLine="off"'));
    expect(extractSdtText(out)).toMatch(/w:multiLine="0"/);
  });
});
