/**
 * `<w:tblLook>` — individual-attribute parsing bug with XMLParser coercion.
 *
 * Per ECMA-376 §17.4.57 CT_TblLook has both a `w:val` attribute (hex
 * bitmask) and six individual CT_OnOff attributes: firstRow, lastRow,
 * firstColumn, lastColumn, noHBand, noVBand.
 *
 * Word often emits ONLY the individual attributes (without `w:val`), in
 * which case the parser reconstructs the hex bitmask from each set flag.
 * But the attribute comparison uses strict string equality:
 *   `if (look['@_w:firstRow'] === '1') value |= 0x0020;`
 *
 * Per `project_xmlparser_numeric_attrs.md`: XMLParser's default
 * `parseAttributeValue: true` coerces `"1"` to the number `1` — so
 * `=== '1'` fails, and the parser wrongly treats every individually-set
 * flag as absent, producing `tblLook="0000"` instead of the real value.
 *
 * This impacts any document where Word saved the tblLook flags in the
 * expanded individual-attribute form — common in recent Word versions
 * and in Open XML SDK-authored documents.
 */

import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('<w:tblLook> individual-attribute parsing', () => {
  it('correctly parses individual-attribute tblLook (no w:val) with XMLParser coercion', async () => {
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
    // Word-style tblLook with individual attributes, NO w:val.
    zipHandler.addFile(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>cell</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>doc</w:t></w:r></w:p>
  </w:body>
</w:document>`
    );
    const buffer = await zipHandler.toBuffer();

    const doc = await Document.loadFromBuffer(buffer);
    const table = doc.getTables()[0] as Table;
    const tblLook = table.getFormatting().tblLook;

    // firstRow (0x0020 = 32) | firstColumn (0x0080 = 128) | noVBand (0x0400 = 1024)
    // = 1184 = 0x04A0
    // Previously this parser returned "0000" because '1' === '1' failed when
    // XMLParser coerced "1" to number 1.
    expect(tblLook).toBe('04A0');
    doc.dispose();
  });

  it('still parses w:val hex string format (regression guard)', async () => {
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
        <w:tblLook w:val="04A0"/>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>cell</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>doc</w:t></w:r></w:p>
  </w:body>
</w:document>`
    );
    const buffer = await zipHandler.toBuffer();
    const doc = await Document.loadFromBuffer(buffer);
    const table = doc.getTables()[0] as Table;
    expect(table.getFormatting().tblLook).toBe('04A0');
    doc.dispose();
  });
});
