/**
 * Style-level table alignment — ST_JcTable `start` / `end` values.
 *
 * Per ECMA-376 Part 1 §17.18.45 ST_JcTable has five values:
 *   start, end, center, left, right.
 *
 * The `TableAlignment` type supports all five, and the main table parser
 * accepts them — but the style-level `parseTableFormattingFromXml`'s
 * whitelist only accepted `left` / `center` / `right`. A table style
 * specifying `<w:jc w:val="start"/>` (the bidi-aware default for modern
 * authors) silently lost its alignment on parse.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithTableStyleJc(jcVal: string): Promise<Buffer> {
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
  <w:style w:type="table" w:styleId="TblAlignTest">
    <w:name w:val="TblAlignTest"/>
    <w:tblPr><w:jc w:val="${jcVal}"/></w:tblPr>
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

function getTableFmt(doc: Document) {
  return doc.getStylesManager().getStyle('TblAlignTest')?.getProperties().tableStyle?.table;
}

describe('Style-level table alignment (§17.18.45 ST_JcTable)', () => {
  const VALUES = ['start', 'end', 'center', 'left', 'right'] as const;

  for (const jcVal of VALUES) {
    it(`parses <w:jc w:val="${jcVal}"/> on a table style`, async () => {
      const buffer = await makeDocxWithTableStyleJc(jcVal);
      const doc = await Document.loadFromBuffer(buffer);
      expect(getTableFmt(doc)?.alignment).toBe(jcVal);
      doc.dispose();
    });
  }
});
