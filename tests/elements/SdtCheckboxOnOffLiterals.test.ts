/**
 * SDT `<w14:checkbox>` — honour every ST_OnOff literal for
 * `<w14:checked w14:val="…"/>`.
 *
 * Per the Word 2010+ w14 extension schema, `<w14:checked>` is a CT_OnOff
 * element in the `w14:` namespace, so its `w14:val` attribute follows
 * ST_OnOff ("true"/"false"/"1"/"0"/"on"/"off") and a bare self-closing
 * element means "on" (true).
 *
 * The SDT checkbox parser used a bespoke comparison:
 *   `checkedVal === 1 || checkedVal === '1' ||
 *    checkedVal === true || checkedVal === 'true'`
 *
 * That silently mishandled `val="on"` (emitted by some Word dialogs and
 * third-party editors) and refused to treat a bare `<w14:checked/>` as
 * true. Round-trip of an `"on"`-authored SDT checkbox came back
 * unchecked.
 *
 * Iteration 94 routes the parse through `parseOoxmlBoolean` with the
 * `@_w14:val` attribute override.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { StructuredDocumentTag } from '../../src/elements/StructuredDocumentTag';

async function loadSdtCheckedValue(valXml: string) {
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
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:sdt>
      <w:sdtPr>
        <w:id w:val="100"/>
        <w14:checkbox>
          ${valXml}
          <w14:checkedState w14:val="2612" w14:font="MS Gothic"/>
          <w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>
        </w14:checkbox>
      </w:sdtPr>
      <w:sdtContent>
        <w:p><w:r><w:t>&#9744;</w:t></w:r></w:p>
      </w:sdtContent>
    </w:sdt>
    <w:p><w:r><w:t>after</w:t></w:r></w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const sdt = doc
    .getBodyElements()
    .find((el): el is StructuredDocumentTag => el instanceof StructuredDocumentTag);
  const checked = sdt?.getCheckboxProperties()?.checked;
  doc.dispose();
  return checked;
}

describe('SDT <w14:checkbox> <w14:checked> — ST_OnOff literal coverage', () => {
  it('val="1" → checked=true (baseline)', async () => {
    expect(await loadSdtCheckedValue('<w14:checked w14:val="1"/>')).toBe(true);
  });
  it('val="0" → checked=false (baseline)', async () => {
    expect(await loadSdtCheckedValue('<w14:checked w14:val="0"/>')).toBe(false);
  });
  it('val="true" → checked=true', async () => {
    expect(await loadSdtCheckedValue('<w14:checked w14:val="true"/>')).toBe(true);
  });
  it('val="false" → checked=false', async () => {
    expect(await loadSdtCheckedValue('<w14:checked w14:val="false"/>')).toBe(false);
  });
  it('val="on" → checked=true', async () => {
    expect(await loadSdtCheckedValue('<w14:checked w14:val="on"/>')).toBe(true);
  });
  it('val="off" → checked=false', async () => {
    expect(await loadSdtCheckedValue('<w14:checked w14:val="off"/>')).toBe(false);
  });
  it('bare <w14:checked/> → checked=true', async () => {
    expect(await loadSdtCheckedValue('<w14:checked/>')).toBe(true);
  });
});
