/**
 * CellVerticalAlignment — ST_VerticalJc `both` value round-trip.
 *
 * Per ECMA-376 Part 1 §17.18.101 ST_VerticalJc has four values:
 * `top`, `center`, `both`, `bottom`. The `CellVerticalAlignment` type
 * omitted `both` (3 values only), and the style-level tcPr parser's
 * whitelist enforced the same narrow set — a table-cell style using
 * `<w:vAlign w:val="both"/>` (distributes the content vertically with
 * equal top/bottom gaps — similar to justified text on the vertical
 * axis) silently lost the value on parse.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithCellVAlign(val: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();
  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );
  zipHandler.addFile(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="table" w:styleId="CellVAlignTest">
    <w:name w:val="CellVAlignTest"/>
    <w:tcPr><w:vAlign w:val="${val}"/></w:tcPr>
  </w:style>
</w:styles>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>test</w:t></w:r></w:p></w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

function getCellFmt(doc: Document) {
  return doc.getStylesManager().getStyle('CellVAlignTest')?.getProperties().tableStyle?.cell;
}

describe('CellVerticalAlignment — ST_VerticalJc (§17.18.101)', () => {
  const VALUES = ['top', 'center', 'both', 'bottom'] as const;
  for (const val of VALUES) {
    it(`parses <w:vAlign w:val="${val}"/> on a table-cell style`, async () => {
      const buffer = await makeDocxWithCellVAlign(val);
      const doc = await Document.loadFromBuffer(buffer);
      expect(getCellFmt(doc)?.verticalAlignment).toBe(val);
      doc.dispose();
    });
  }
});
