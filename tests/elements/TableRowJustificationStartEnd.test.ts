/**
 * TableRow `<w:jc>` — ST_JcTable parser acceptance and save-time mapping.
 *
 * ISO/IEC 29500 §17.18.45 ST_JcTable defines five values for the row-level
 * `<w:jc>` attribute: `start`, `center`, `end`, `left`, `right`. The first
 * two are the bidi-aware spellings (they flip under RTL).
 *
 * HOWEVER, the Open XML SDK's row-level `TableRowAlignmentValues` enum
 * (which the strict OOXML validator uses) accepts only three values:
 * `center` / `left` / `right`. This is narrower than ISO/ECMA. Emitting
 * `<w:jc w:val="start"/>` inside `<w:trPr>` therefore fails validation
 * with "The attribute 'w:val' has invalid value 'start'. The Enumeration
 * constraint failed." — even though the ECMA spec allows it.
 *
 * The framework reconciles this by:
 *   1. Parsing all five values faithfully (so RTL-aware documents load
 *      without information loss).
 *   2. Mapping `start` → `left` and `end` → `right` on save so the
 *      emitted XML passes SDK validation.
 *
 * The mapping IS lossy: an RTL-section row authored with `start` will
 * lose its RTL-awareness on round-trip. There is no fix for that without
 * changing the SDK / validator. See `project_sdk_rowjc_narrower_than_spec`
 * memory for detail.
 */

import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithRowJc(jcVal: string): Promise<Buffer> {
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
        <w:trPr><w:jc w:val="${jcVal}"/></w:trPr>
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
  return await zipHandler.toBuffer();
}

describe('TableRow w:jc parser — all ST_JcTable values accepted (§17.18.45)', () => {
  const VALUES = ['start', 'end', 'center', 'left', 'right'] as const;

  for (const jcVal of VALUES) {
    it(`parses row-level <w:jc w:val="${jcVal}"/>`, async () => {
      const buffer = await makeDocxWithRowJc(jcVal);
      const doc = await Document.loadFromBuffer(buffer);
      const table = doc.getTables()[0] as Table;
      const row = table.getRows()[0]!;
      expect(row.getJustification()).toBe(jcVal);
      doc.dispose();
    });
  }
});

describe('TableRow w:jc save-time mapping — bidi-aware → LTR (SDK constraint)', () => {
  // The SDK's TableRowAlignmentValues accepts only Center/Left/Right. Emitting
  // start/end at row level fails strict validation. The framework maps them
  // on save so output stays validator-clean.
  it('maps "start" → "left" on save (validator compatibility)', async () => {
    const buffer = await makeDocxWithRowJc('start');
    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const docFile = zip.getFile('word/document.xml');
    const content = docFile?.content;
    const xml = content instanceof Buffer ? content.toString('utf8') : String(content);

    expect(xml).toMatch(/<w:trPr>[\s\S]*?<w:jc\s+w:val="left"[\s\S]*?<\/w:trPr>/);
    expect(xml).not.toMatch(/<w:trPr>[\s\S]*?<w:jc\s+w:val="start"/);
  });

  it('maps "end" → "right" on save (validator compatibility)', async () => {
    const buffer = await makeDocxWithRowJc('end');
    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const docFile = zip.getFile('word/document.xml');
    const content = docFile?.content;
    const xml = content instanceof Buffer ? content.toString('utf8') : String(content);

    expect(xml).toMatch(/<w:trPr>[\s\S]*?<w:jc\s+w:val="right"[\s\S]*?<\/w:trPr>/);
    expect(xml).not.toMatch(/<w:trPr>[\s\S]*?<w:jc\s+w:val="end"/);
  });

  it('passes through "center" / "left" / "right" unchanged', async () => {
    for (const val of ['center', 'left', 'right'] as const) {
      const buffer = await makeDocxWithRowJc(val);
      const doc = await Document.loadFromBuffer(buffer);
      const out = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const docFile = zip.getFile('word/document.xml');
      const content = docFile?.content;
      const xml = content instanceof Buffer ? content.toString('utf8') : String(content);

      const trPrRe = new RegExp(`<w:trPr>[\\s\\S]*?<w:jc\\s+w:val="${val}"[\\s\\S]*?<\\/w:trPr>`);
      expect(xml).toMatch(trPrRe);
    }
  });
});
