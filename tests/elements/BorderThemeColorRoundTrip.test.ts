/**
 * CT_Border extended attribute round-trip — themeColor / themeTint /
 * themeShade / shadow / frame.
 *
 * Per ECMA-376 Part 1 §17.18.2 CT_Border has the following attributes:
 *   val (required), sz, space, color, themeColor, themeTint, themeShade,
 *   shadow, frame.
 *
 * The framework's BorderDefinition / BorderProperties interfaces only
 * declared the first four (val/sz/space/color) — parser silently dropped
 * the rest, and generator emission paths only knew about the "basic"
 * quartet. Documents authored by Word with themed borders or shadow /
 * frame flags lost all of that metadata on load → save round-trip.
 *
 * This iteration widens both interfaces and extends
 * `XMLBuilder.createBorder` (used by the main Table and TableCell emit
 * paths) and `DocumentParser.parseBordersFromXml` (used by every
 * style-level / table-level border parse) to handle the full attribute
 * set.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('CT_Border extended attributes — themeColor / shadow / frame round-trip', () => {
  describe('XMLBuilder.createBorder unit shape', () => {
    it('emits w:themeColor / w:themeTint / w:themeShade when set', () => {
      const el = XMLBuilder.createBorder('top', {
        style: 'single',
        size: 4,
        color: 'auto',
        themeColor: 'accent1',
        themeTint: '66',
        themeShade: '80',
      });
      const xml = XMLBuilder.elementToString(el);
      expect(xml).toMatch(/w:val="single"/);
      expect(xml).toMatch(/w:themeColor="accent1"/);
      expect(xml).toMatch(/w:themeTint="66"/);
      expect(xml).toMatch(/w:themeShade="80"/);
    });

    it('emits w:shadow and w:frame as 1/0', () => {
      const el = XMLBuilder.createBorder('bottom', {
        style: 'single',
        size: 4,
        shadow: true,
        frame: false,
      });
      const xml = XMLBuilder.elementToString(el);
      expect(xml).toMatch(/w:shadow="1"/);
      expect(xml).toMatch(/w:frame="0"/);
    });

    it('omits the extended attrs when undefined', () => {
      const el = XMLBuilder.createBorder('right', { style: 'single', size: 4 });
      const xml = XMLBuilder.elementToString(el);
      expect(xml).not.toMatch(/w:themeColor=|w:themeTint=|w:themeShade=|w:shadow=|w:frame=/);
    });
  });

  describe('parser round-trip via table cell borders', () => {
    it('preserves themeColor / themeTint / themeShade on table cell borders', async () => {
      // Author a DOCX with a themed border inside a table cell.
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
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="5000" w:type="pct"/>
            <w:tcBorders>
              <w:top w:val="single" w:sz="4" w:color="auto" w:themeColor="accent1" w:themeTint="66" w:themeShade="80"/>
            </w:tcBorders>
          </w:tcPr>
          <w:p><w:r><w:t>cell</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>doc</w:t></w:r></w:p>
  </w:body>
</w:document>`
      );
      const buffer = await zipHandler.toBuffer();

      const doc = await Document.loadFromBuffer(buffer);
      const out = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const docFile = zip.getFile('word/document.xml');
      const content = docFile?.content;
      const xml = content instanceof Buffer ? content.toString('utf8') : String(content);

      // The re-emitted top cell border must still carry themeColor / themeTint /
      // themeShade. Previously all three were silently stripped.
      expect(xml).toMatch(/<w:top[^>]*w:themeColor="accent1"/);
      expect(xml).toMatch(/<w:top[^>]*w:themeTint="66"/);
      expect(xml).toMatch(/<w:top[^>]*w:themeShade="80"/);
    });
  });

  describe('Paragraph pBdr direct emission', () => {
    it('emits themeColor / shadow / frame on paragraph borders', () => {
      const p = new Paragraph();
      p.addText('x');
      p.setBorder({
        top: {
          style: 'single',
          size: 4,
          color: 'auto',
          themeColor: 'accent1',
          themeTint: '99',
          shadow: true,
          frame: false,
        },
      });
      const xml = XMLBuilder.elementToString(p.toXML());
      expect(xml).toMatch(/<w:pBdr>[\s\S]*<w:top[^>]*w:themeColor="accent1"/);
      expect(xml).toMatch(/<w:top[^>]*w:themeTint="99"/);
      expect(xml).toMatch(/<w:top[^>]*w:shadow="1"/);
      expect(xml).toMatch(/<w:top[^>]*w:frame="0"/);
    });
  });

  describe('Run character border (w:bdr) emission', () => {
    it('emits themeColor on character borders', () => {
      const run = new Run('x', {
        border: { style: 'single', size: 4, themeColor: 'accent2', themeShade: '80' },
      });
      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toMatch(/<w:bdr[^>]*w:themeColor="accent2"/);
      expect(xml).toMatch(/<w:bdr[^>]*w:themeShade="80"/);
    });
  });
});
