/**
 * Style-level pPr CT_OnOff flags — extended coverage beyond iteration-25's
 * four-flag set (keepNext / keepLines / pageBreakBefore / contextualSpacing).
 *
 * CT_PPrBase per ECMA-376 Part 1 §17.3.1 has ~17 CT_OnOff children. The main
 * paragraph parser (`parseParagraphProperties`) handles all of them; the
 * style-level parser (`parseParagraphFormattingFromXml`) previously handled
 * only the four iteration-25 flags and silently dropped the rest.
 *
 * This suite guards against regressions for 13 additional flags:
 *   - widowControl      §17.3.1.40
 *   - suppressLineNumbers §17.3.1.34
 *   - bidi               §17.3.1.6
 *   - mirrorIndents      §17.3.1.18
 *   - adjustRightInd     §17.3.1.1
 *   - suppressAutoHyphens §17.3.1.33
 *   - kinsoku            §17.3.1.16
 *   - wordWrap           §17.3.1.45
 *   - overflowPunct      §17.3.1.24
 *   - topLinePunct       §17.3.1.43
 *   - autoSpaceDE        §17.3.1.2
 *   - autoSpaceDN        §17.3.1.3
 *   - suppressOverlap    §17.3.1.34
 *
 * Parser cases: each flag × six ST_OnOff literal forms (0/false/off/bare/1/on).
 * Generator cases: emit explicit-false via w:val="0" (NOT bare element, which
 * would mean true).
 */

import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

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
  <w:style w:type="paragraph" w:styleId="CtOnOffTestExt">
    <w:name w:val="CtOnOffTestExt"/>
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
  return doc.getStylesManager().getStyle('CtOnOffTestExt')?.getParagraphFormatting();
}

const EXTENDED_FLAGS = [
  'widowControl',
  'suppressLineNumbers',
  'bidi',
  'mirrorIndents',
  'adjustRightInd',
  'suppressAutoHyphens',
  'kinsoku',
  'wordWrap',
  'overflowPunct',
  'topLinePunct',
  'autoSpaceDE',
  'autoSpaceDN',
  'suppressOverlap',
] as const;

describe('Style pPr CT_OnOff — extended flags honour every ST_OnOff literal', () => {
  for (const flag of EXTENDED_FLAGS) {
    it(`<w:${flag} w:val="0"/> parses as false (not true, not dropped)`, async () => {
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

/**
 * Generator-side coverage: Style.toXML() must emit each flag's
 * explicit-false as `<w:{flag} w:val="0"/>`, not omit it entirely.
 * Omission means "inherit from base style" per ST_OnOff semantics.
 * Without this, a style that explicitly overrides a based-on style's
 * enabled flag silently reverts to the inherited true on save.
 *
 * Calling toXML() directly exercises the generator regardless of the raw-XML
 * passthrough cache (`mergeStylesWithOriginal` would otherwise re-emit the
 * original bytes verbatim for an unmodified style).
 */
describe('Style pPr CT_OnOff — extended flags survive round-trip via toXML()', () => {
  for (const flag of EXTENDED_FLAGS) {
    it(`emits <w:${flag} w:val="0"/> for explicit false`, () => {
      const style = new Style({
        styleId: 'CtOnOffTestExt',
        type: 'paragraph',
        name: 'CtOnOffTestExt',
        paragraphFormatting: { [flag]: false },
      });
      const xml = XMLBuilder.elementToString(style.toXML());
      expect(xml).toContain(`<w:${flag} w:val="0"/>`);
    });

    it(`emits bare <w:${flag}/> for explicit true`, () => {
      const style = new Style({
        styleId: 'CtOnOffTestExt',
        type: 'paragraph',
        name: 'CtOnOffTestExt',
        paragraphFormatting: { [flag]: true },
      });
      const xml = XMLBuilder.elementToString(style.toXML());
      expect(xml).toMatch(new RegExp(`<w:${flag}(/>|\\s+w:val="1"\\s*/>)`));
    });

    it(`omits <w:${flag}/> entirely when undefined`, () => {
      const style = new Style({
        styleId: 'CtOnOffTestExt',
        type: 'paragraph',
        name: 'CtOnOffTestExt',
        paragraphFormatting: {}, // flag absent
      });
      const xml = XMLBuilder.elementToString(style.toXML());
      expect(xml).not.toContain(`<w:${flag}`);
    });
  }
});
