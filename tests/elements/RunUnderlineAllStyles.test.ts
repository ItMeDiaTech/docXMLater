/**
 * Run underline — all 18 ST_Underline values per ECMA-376 §17.3.2.40.
 *
 * The `RunFormatting.underline` type was declared as:
 *   boolean | 'single' | 'double' | 'thick' | 'dotted' | 'dash' | 'none'
 *
 * …which covers only 6 of the 18 ST_Underline enum values. The style-level
 * rPr parser's whitelist (DocumentParser.ts:~9453) enforced the same
 * narrow set, falling back to `underline = true` for `words`,
 * `dottedHeavy`, `dashedHeavy`, `dashLong`, `dashLongHeavy`, `dotDash`,
 * `dashDotHeavy`, `dotDotDash`, `dashDotDotHeavy`, `wave`, `wavyHeavy`,
 * and `wavyDouble` — a character style using any of those 12 values lost
 * the specific style on parse.
 *
 * ST_Underline values (18 total):
 *   single words double thick dotted dottedHeavy dash dashedHeavy
 *   dashLong dashLongHeavy dotDash dashDotHeavy dotDotDash
 *   dashDotDotHeavy wave wavyHeavy wavyDouble none
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyleUnderline(uVal: string): Promise<Buffer> {
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
  <w:style w:type="character" w:styleId="UTest">
    <w:name w:val="UTest"/>
    <w:rPr><w:u w:val="${uVal}"/></w:rPr>
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

const ALL_UNDERLINE_VALUES = [
  'single',
  'words',
  'double',
  'thick',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashedHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dashDotHeavy',
  'dotDotDash',
  'dashDotDotHeavy',
  'wave',
  'wavyHeavy',
  'wavyDouble',
  'none',
] as const;

describe('Style rPr — underline (w:u w:val, §17.3.2.40)', () => {
  for (const uVal of ALL_UNDERLINE_VALUES) {
    it(`parses <w:u w:val="${uVal}"/> as underline: "${uVal}"`, async () => {
      const buffer = await makeDocxWithStyleUnderline(uVal);
      const doc = await Document.loadFromBuffer(buffer);
      const rPr = doc.getStylesManager().getStyle('UTest')?.getRunFormatting();
      expect(rPr?.underline).toBe(uVal);
      doc.dispose();
    });
  }

  it('falls back to boolean true for an unknown / out-of-spec value', async () => {
    const buffer = await makeDocxWithStyleUnderline('bogusStyle');
    const doc = await Document.loadFromBuffer(buffer);
    const rPr = doc.getStylesManager().getStyle('UTest')?.getRunFormatting();
    // Unknown values still map to `true` (underline enabled) rather than
    // being dropped entirely — matches the main parser's fallback.
    expect(rPr?.underline).toBe(true);
    doc.dispose();
  });
});
