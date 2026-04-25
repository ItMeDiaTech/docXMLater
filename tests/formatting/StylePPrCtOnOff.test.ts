/**
 * Style-level pPr CT_OnOff flags — `w:val` handling.
 *
 * `DocumentParser.parseParagraphFormattingFromXml` detects four pPr
 * CT_OnOff flags via substring match:
 *
 *   if (pPrXml.includes('<w:keepNext/>') || pPrXml.includes('<w:keepNext ')) {
 *     formatting.keepNext = true;
 *   }
 *
 * Bug this suite guards against:
 *   - `<w:keepNext w:val="0"/>` matches `<w:keepNext ` and is therefore
 *     hard-coded to `true`, silently flipping an explicit-false override
 *     into an enabled flag. Same pattern for `w:keepLines`,
 *     `w:pageBreakBefore`, and `w:contextualSpacing`. Any style that
 *     used the explicit-false form to override a based-on style's
 *     enabled CT_OnOff was silently inverted on load → save.
 *
 * This is the style-level equivalent of the CT_OnOff audit done in
 * earlier iterations for pPr / rPr / settings / section / table cells.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStylePPr(pPrInner: string): Promise<Buffer> {
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
  <w:style w:type="paragraph" w:styleId="CtOnOffTest">
    <w:name w:val="CtOnOffTest"/>
    <w:pPr>${pPrInner}</w:pPr>
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

function getFmt(doc: Document) {
  return doc.getStylesManager().getStyle('CtOnOffTest')?.getParagraphFormatting();
}

const FLAGS = ['keepNext', 'keepLines', 'pageBreakBefore', 'contextualSpacing'] as const;

describe('Style pPr CT_OnOff — honours w:val="0" (explicit false)', () => {
  for (const flag of FLAGS) {
    it(`<w:${flag} w:val="0"/> parses as false (not true)`, async () => {
      const buffer = await makeDocxWithStylePPr(`<w:${flag} w:val="0"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const fmt = getFmt(doc);
      expect(fmt?.[flag]).toBe(false);
      doc.dispose();
    });

    it(`<w:${flag} w:val="false"/> parses as false`, async () => {
      const buffer = await makeDocxWithStylePPr(`<w:${flag} w:val="false"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const fmt = getFmt(doc);
      expect(fmt?.[flag]).toBe(false);
      doc.dispose();
    });

    it(`<w:${flag} w:val="off"/> parses as false`, async () => {
      const buffer = await makeDocxWithStylePPr(`<w:${flag} w:val="off"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const fmt = getFmt(doc);
      expect(fmt?.[flag]).toBe(false);
      doc.dispose();
    });

    it(`<w:${flag}/> (bare) parses as true`, async () => {
      const buffer = await makeDocxWithStylePPr(`<w:${flag}/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const fmt = getFmt(doc);
      expect(fmt?.[flag]).toBe(true);
      doc.dispose();
    });

    it(`<w:${flag} w:val="1"/> parses as true`, async () => {
      const buffer = await makeDocxWithStylePPr(`<w:${flag} w:val="1"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const fmt = getFmt(doc);
      expect(fmt?.[flag]).toBe(true);
      doc.dispose();
    });

    it(`<w:${flag} w:val="on"/> parses as true`, async () => {
      const buffer = await makeDocxWithStylePPr(`<w:${flag} w:val="on"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const fmt = getFmt(doc);
      expect(fmt?.[flag]).toBe(true);
      doc.dispose();
    });
  }
});
