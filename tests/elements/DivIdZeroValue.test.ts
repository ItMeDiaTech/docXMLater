/**
 * `<w:divId w:val="0"/>` — HTML div reference with the zero index
 * must survive PARSING with its explicit zero value preserved.
 *
 * Per ECMA-376 Part 1 §17.3.1.10 (CT_DivId, paragraph context) and
 * §17.4.9 (CT_DivId, row context), `w:val` is ST_DecimalNumber —
 * `xsd:integer` — so 0 is a valid value (referencing the first div in
 * the web-settings part).
 *
 * Two compounding bugs silently dropped zero-valued divId on load:
 *   1. XMLParser's `parseAttributeValue: true` coerced `"0"` to the
 *      number `0`.
 *   2. The paragraph parser (`DocumentParser.ts:2577`) used a truthy
 *      check (`if (divIdVal)`); the row parser (`DocumentParser.ts:6958`)
 *      used `val > 0`. Both dropped the coerced numeric zero and never
 *      called `setDivId(0)`.
 *
 * Both emitters (`Paragraph.ts:3460`, `TableRow.ts` trPr builder) use
 * `!== undefined`, so the parser/emitter asymmetry broke round-trip
 * for every divId=0 reference.
 *
 * Iteration 106 swaps the truthy / `> 0` gates for
 * `isExplicitlySet` + `safeParseInt` on both parsers. The tests here
 * assert on the parsed in-memory state rather than on a save round
 * trip so we don't need a fully-declared web-settings part (the SDK
 * validator rejects divId references to div IDs that aren't declared
 * in webSettings.xml).
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadAndReadParagraphDivId(valAttr: string): Promise<number | undefined> {
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
    <w:p>
      <w:pPr>
        <w:divId w:val="${valAttr}"/>
      </w:pPr>
      <w:r><w:t>x</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const paragraph = doc.getParagraphs()[0]!;
  const divId = paragraph.getFormatting().divId;
  doc.dispose();
  return divId;
}

async function loadAndReadRowDivId(valAttr: string): Promise<number | undefined> {
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
      <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:trPr>
          <w:divId w:val="${valAttr}"/>
        </w:trPr>
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
  const doc = await Document.loadFromBuffer(buffer);
  const row = doc.getTables()[0]!.getRows()[0]!;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const divId = (row as any).getDivId?.();
  doc.dispose();
  return divId;
}

describe('<w:divId w:val="0"/> parses to the literal number 0', () => {
  it('preserves val="0" on a paragraph (§17.3.1.10)', async () => {
    expect(await loadAndReadParagraphDivId('0')).toBe(0);
  });

  it('preserves val="42" on a paragraph (regression guard)', async () => {
    expect(await loadAndReadParagraphDivId('42')).toBe(42);
  });

  it('preserves val="0" on a table row (§17.4.9)', async () => {
    expect(await loadAndReadRowDivId('0')).toBe(0);
  });

  it('preserves val="42" on a table row (regression guard)', async () => {
    expect(await loadAndReadRowDivId('42')).toBe(42);
  });
});
