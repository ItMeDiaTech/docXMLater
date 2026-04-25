/**
 * Style-level rPr CT_OnOff flags — `w:val` handling.
 *
 * `DocumentParser.parseRunFormattingFromXml` (the style-level rPr parser
 * at ~9336) detects CT_OnOff flags via substring-include:
 *
 *   if (rPrXml.includes('<w:b/>') || rPrXml.includes('<w:b ')) {
 *     formatting.bold = true;
 *   }
 *
 * Same bug as was fixed on the pPr side in iteration 25:
 *   - `<w:b w:val="0"/>` matches `<w:b ` and is therefore hard-coded to
 *     `true`, silently flipping an explicit-false override into an
 *     enabled flag. Same pattern for `w:i`, `w:strike`, `w:smallCaps`,
 *     `w:caps`. Any character style that used the explicit-false form to
 *     override a based-on style's enabled CT_OnOff was silently inverted
 *     on load → save.
 *
 * Also, the parser never sets these flags to `false` at all — it only
 * writes `true`. So explicit-false had nowhere to land even if detected.
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
  <w:style w:type="character" w:styleId="RPrCtOnOffTest">
    <w:name w:val="RPrCtOnOffTest"/>
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
  return doc.getStylesManager().getStyle('RPrCtOnOffTest')?.getRunFormatting();
}

const FLAGS = [
  { tag: 'b', field: 'bold' },
  { tag: 'i', field: 'italic' },
  { tag: 'strike', field: 'strike' },
  { tag: 'smallCaps', field: 'smallCaps' },
  { tag: 'caps', field: 'allCaps' },
] as const;

describe('Style rPr CT_OnOff — honours w:val="0" (explicit false)', () => {
  for (const { tag, field } of FLAGS) {
    it(`<w:${tag} w:val="0"/> parses as false (not true)`, async () => {
      const buffer = await makeDocxWithStyleRPr(`<w:${tag} w:val="0"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const rPr = getRPr(doc);
      expect(rPr?.[field]).toBe(false);
      doc.dispose();
    });

    it(`<w:${tag} w:val="false"/> parses as false`, async () => {
      const buffer = await makeDocxWithStyleRPr(`<w:${tag} w:val="false"/>`);
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

    it(`<w:${tag} w:val="1"/> parses as true`, async () => {
      const buffer = await makeDocxWithStyleRPr(`<w:${tag} w:val="1"/>`);
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
