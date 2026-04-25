/**
 * `<w:tc><w:tcPr><w:vAlign w:val="both"/></w:tcPr>` — ST_VerticalJc
 * "both" round-trip.
 *
 * Per ECMA-376 Part 1 §17.18.101 ST_VerticalJc, the enumeration for
 * cell-level vertical alignment has FOUR values: top, center, both,
 * bottom. ("both" = justified vertical alignment — content is
 * distributed from top to bottom of the cell.)
 *
 * The main table cell parser (`DocumentParser.ts::parseTableCellFromObject`
 * around line 7158) whitelisted only three values (top / center /
 * bottom), silently dropping `"both"` on load. The style-level cell
 * parser (`parseTableCellFormattingFromXml`) already accepts all four,
 * so the asymmetry meant Word's "Justify vertically" cell alignment
 * round-tripped correctly *via style inheritance* but not *as a direct
 * override*.
 *
 * Iteration 102 extends the main parser whitelist to include "both".
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { Table } from '../../src/elements/Table';

async function loadAndReadFirstCellVAlign(valAttr: string): Promise<string | undefined> {
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
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="5000" w:type="pct"/>
            <w:vAlign w:val="${valAttr}"/>
          </w:tcPr>
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
  const table = doc.getTables()[0] as Table;
  const cell = table.getRows()[0]!.getCells()[0]!;
  const vAlign = cell.getVerticalAlignment();
  doc.dispose();
  return vAlign;
}

describe('<w:vAlign> ST_VerticalJc "both" — main cell parser', () => {
  it('preserves val="both" (previously silently dropped)', async () => {
    expect(await loadAndReadFirstCellVAlign('both')).toBe('both');
  });

  it('preserves val="top" (regression guard)', async () => {
    expect(await loadAndReadFirstCellVAlign('top')).toBe('top');
  });

  it('preserves val="center" (regression guard)', async () => {
    expect(await loadAndReadFirstCellVAlign('center')).toBe('center');
  });

  it('preserves val="bottom" (regression guard)', async () => {
    expect(await loadAndReadFirstCellVAlign('bottom')).toBe('bottom');
  });

  it('drops unknown values (regression guard)', async () => {
    // Unknown ST_VerticalJc values should fall through to undefined,
    // matching the prior whitelist behaviour for invalid input.
    expect(await loadAndReadFirstCellVAlign('invalid')).toBeUndefined();
  });
});
