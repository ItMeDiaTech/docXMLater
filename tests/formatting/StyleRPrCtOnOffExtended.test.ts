/**
 * Style-level rPr CT_OnOff flags — extended coverage beyond iteration-40's
 * five-flag set (b / i / strike / smallCaps / caps).
 *
 * CT_RPr per ECMA-376 §17.3.2 has ~15 CT_OnOff children. The style-level
 * rPr parser (`parseRunFormattingFromXml`) previously ignored most of
 * them entirely — the substring-include scan only covered the five
 * iteration-40 flags. Any character style that set `<w:dstrike>`,
 * `<w:outline>`, `<w:vanish>`, `<w:rtl>`, etc. was silently dropped on
 * parse, so a programmatic resave (bypassing raw-XML passthrough) lost
 * the style override.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyleRPr(rPrInner: string): Promise<Buffer> {
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
  <w:style w:type="character" w:styleId="RPrExt">
    <w:name w:val="RPrExt"/>
    <w:rPr>${rPrInner}</w:rPr>
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

function getRPr(doc: Document) {
  return doc.getStylesManager().getStyle('RPrExt')?.getRunFormatting();
}

// Pairs of (XML tag, RunFormatting field name).
const EXTENDED_FLAGS = [
  { tag: 'bCs', field: 'complexScriptBold' },
  { tag: 'iCs', field: 'complexScriptItalic' },
  { tag: 'cs', field: 'complexScript' },
  { tag: 'dstrike', field: 'dstrike' },
  { tag: 'outline', field: 'outline' },
  { tag: 'shadow', field: 'shadow' },
  { tag: 'emboss', field: 'emboss' },
  { tag: 'imprint', field: 'imprint' },
  { tag: 'rtl', field: 'rtl' },
  { tag: 'vanish', field: 'vanish' },
  { tag: 'noProof', field: 'noProof' },
  { tag: 'snapToGrid', field: 'snapToGrid' },
  { tag: 'specVanish', field: 'specVanish' },
  { tag: 'webHidden', field: 'webHidden' },
] as const;

describe('Style rPr CT_OnOff — extended flags honour every ST_OnOff literal', () => {
  for (const { tag, field } of EXTENDED_FLAGS) {
    it(`<w:${tag} w:val="0"/> parses as false (not dropped)`, async () => {
      const buffer = await makeDocxWithStyleRPr(`<w:${tag} w:val="0"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const rPr = getRPr(doc);
      expect(rPr?.[field]).toBe(false);
      doc.dispose();
    });

    it(`<w:${tag}/> (bare) parses as true`, async () => {
      const buffer = await makeDocxWithStyleRPr(`<w:${tag}/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const rPr = getRPr(doc);
      expect(rPr?.[field]).toBe(true);
      doc.dispose();
    });

    it(`<w:${tag} w:val="on"/> parses as true`, async () => {
      const buffer = await makeDocxWithStyleRPr(`<w:${tag} w:val="on"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const rPr = getRPr(doc);
      expect(rPr?.[field]).toBe(true);
      doc.dispose();
    });

    it(`absent <w:${tag}> leaves ${field} undefined`, async () => {
      const buffer = await makeDocxWithStyleRPr('');
      const doc = await Document.loadFromBuffer(buffer);
      const rPr = getRPr(doc);
      expect(rPr?.[field]).toBeUndefined();
      doc.dispose();
    });
  }
});
