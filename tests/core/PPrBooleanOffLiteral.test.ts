/**
 * Paragraph CT_OnOff parsers — honour the "off" / "on" ST_OnOff literals.
 *
 * Per ECMA-376 §17.17.4 ST_OnOff, every CT_OnOff element's `w:val` attribute
 * can take one of six literals: "true" / "false" / "1" / "0" / "on" / "off".
 *
 * The main paragraph parser for `<w:widowControl>`, `<w:bidi>`, and
 * `<w:adjustRightInd>` used bespoke comparisons that only checked for
 * "0" / "false" / boolean false / number 0 — missing the "off" literal
 * entirely. A Word-authored paragraph with `<w:widowControl w:val="off"/>`
 * silently flipped to widowControl=TRUE on parse, changing rendering
 * behavior.
 *
 * Iteration 92 routes all three through `parseOoxmlBoolean`, which
 * correctly handles every ST_OnOff literal form.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadParagraph(pPrInner: string) {
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
      <w:pPr>${pPrInner}</w:pPr>
      <w:r><w:t>x</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const paragraph = doc.getParagraphs()[0]!;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const fmt = (paragraph as any).formatting;
  doc.dispose();
  return fmt;
}

describe('<w:widowControl> / <w:bidi> / <w:adjustRightInd> — ST_OnOff literal coverage', () => {
  describe('<w:widowControl>', () => {
    it('val="off" → false', async () => {
      const fmt = await loadParagraph('<w:widowControl w:val="off"/>');
      expect(fmt.widowControl).toBe(false);
    });
    it('val="on" → true', async () => {
      const fmt = await loadParagraph('<w:widowControl w:val="on"/>');
      expect(fmt.widowControl).toBe(true);
    });
    it('val="0" → false', async () => {
      const fmt = await loadParagraph('<w:widowControl w:val="0"/>');
      expect(fmt.widowControl).toBe(false);
    });
    it('bare element → true', async () => {
      const fmt = await loadParagraph('<w:widowControl/>');
      expect(fmt.widowControl).toBe(true);
    });
  });

  describe('<w:bidi>', () => {
    it('val="off" → false', async () => {
      const fmt = await loadParagraph('<w:bidi w:val="off"/>');
      expect(fmt.bidi).toBe(false);
    });
    it('val="on" → true', async () => {
      const fmt = await loadParagraph('<w:bidi w:val="on"/>');
      expect(fmt.bidi).toBe(true);
    });
  });

  describe('<w:adjustRightInd>', () => {
    it('val="off" → false', async () => {
      const fmt = await loadParagraph('<w:adjustRightInd w:val="off"/>');
      expect(fmt.adjustRightInd).toBe(false);
    });
    it('val="on" → true', async () => {
      const fmt = await loadParagraph('<w:adjustRightInd w:val="on"/>');
      expect(fmt.adjustRightInd).toBe(true);
    });
  });
});
