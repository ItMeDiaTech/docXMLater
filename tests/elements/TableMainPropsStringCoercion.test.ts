/**
 * Main-path table property parser — ST_String attributes must
 * round-trip as strings (type-contract safety).
 *
 * Per ECMA-376 Part 1:
 *   - `w:tblStyle w:val` is ST_String (§17.7.4.62) — table style ID
 *   - `w:tblCaption w:val` is ST_String (§17.4.62) — accessible caption
 *   - `w:tblDescription w:val` is ST_String (§17.4.63) — accessible descr
 *
 * Iteration 128 fixed the *previousProperties* variant (inside
 * `w:tblPrChange`) but the main-path parser in
 * `parseTablePropertiesFromObject` still had the same bug:
 *
 *     const styleId = tblPrObj['w:tblStyle']['@_w:val'];
 *     if (styleId) { table.setStyle(styleId); }  // could be number
 *
 *     const caption = tblPrObj['w:tblCaption']['@_w:val'];
 *     if (caption) table.setCaption(caption);      // could be number
 *
 *     const description = tblPrObj['w:tblDescription']['@_w:val'];
 *     if (description) table.setDescription(description);  // could be number
 *
 * All three setters declare `(value: string)`. Storing a number
 * violates the contract and breaks downstream string operations on
 * `table.getStyle() / getCaption() / getDescription()`.
 *
 * Iteration 129 casts all three through `String(...)` at the read site.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadTableWith(tblPrInnerXml: string) {
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        ${tblPrInnerXml}
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>c</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>d</w:t></w:r></w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  return Document.loadFromBuffer(buffer);
}

describe('Main-path table ST_String attribute type-contract preservation', () => {
  it('stores <w:tblStyle w:val="2025"/> as the STRING "2025"', async () => {
    const doc = await loadTableWith('<w:tblStyle w:val="2025"/>');
    const table = doc.getTables()[0]!;
    const style = table.getStyle();
    doc.dispose();
    expect(typeof style).toBe('string');
    expect(style).toBe('2025');
  });

  it('stores <w:tblStyle w:val="1"/> as the STRING "1"', async () => {
    const doc = await loadTableWith('<w:tblStyle w:val="1"/>');
    const table = doc.getTables()[0]!;
    const style = table.getStyle();
    doc.dispose();
    expect(typeof style).toBe('string');
    expect(style).toBe('1');
  });

  it('stores <w:tblCaption w:val="42"/> as the STRING "42"', async () => {
    const doc = await loadTableWith('<w:tblCaption w:val="42"/>');
    const table = doc.getTables()[0]!;
    const caption = table.getCaption();
    doc.dispose();
    expect(typeof caption).toBe('string');
    expect(caption).toBe('42');
  });

  it('stores <w:tblDescription w:val="99"/> as the STRING "99"', async () => {
    const doc = await loadTableWith('<w:tblDescription w:val="99"/>');
    const table = doc.getTables()[0]!;
    const description = table.getDescription();
    doc.dispose();
    expect(typeof description).toBe('string');
    expect(description).toBe('99');
  });

  it('string methods are callable on parsed numeric tblStyle', async () => {
    const doc = await loadTableWith('<w:tblStyle w:val="2025"/>');
    const table = doc.getTables()[0]!;
    const style = table.getStyle();
    doc.dispose();
    // Pre-fix: style was number 2025, .startsWith would throw.
    expect(() => style?.startsWith('2')).not.toThrow();
    expect(style?.startsWith('20')).toBe(true);
  });

  it('preserves non-numeric tblStyle "TableGrid" (regression guard)', async () => {
    const doc = await loadTableWith('<w:tblStyle w:val="TableGrid"/>');
    const table = doc.getTables()[0]!;
    const style = table.getStyle();
    doc.dispose();
    expect(style).toBe('TableGrid');
  });

  it('preserves non-numeric caption/description (regression guard)', async () => {
    const doc = await loadTableWith(
      '<w:tblCaption w:val="Quarterly Sales"/><w:tblDescription w:val="FY26 summary"/>'
    );
    const table = doc.getTables()[0]!;
    expect(table.getCaption()).toBe('Quarterly Sales');
    expect(table.getDescription()).toBe('FY26 summary');
    doc.dispose();
  });
});
