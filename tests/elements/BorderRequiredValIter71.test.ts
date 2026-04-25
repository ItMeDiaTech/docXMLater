/**
 * Final `w:val` compliance sites (iteration-71 audit).
 *
 * Per ECMA-376 Part 1 §17.18.2 CT_Border, `w:val` (ST_Border) is REQUIRED.
 *
 * Remaining emission sites with the `if (border.style) attrs['w:val'] = ...`
 * truthy-gate pattern — same bug as iterations 68-70 but on:
 *
 *   1. Run-level `<w:bdr>` (character border, §17.3.2.5) inside `<w:rPr>`.
 *   2. Style-level `<w:tl2br>` (diagonal top-left to bottom-right cell border).
 *   3. Style-level `<w:tr2bl>` (diagonal top-right to bottom-left cell border).
 *
 * All three now default `w:val` to `"nil"` when the consumer set only
 * size / color / space.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('Remaining CT_Border w:val sites (iter-71)', () => {
  describe('Run w:bdr (character border §17.3.2.5)', () => {
    it('defaults w:val to "nil" when run border sets only size/color', () => {
      const run = new Run('x', { border: { size: 4, color: '000000' } });
      const xml = XMLBuilder.elementToString(run.toXML());
      expect(xml).toMatch(/<w:bdr[^>]*w:val="nil"[^>]*w:sz="4"/);
      expect(xml).not.toMatch(/<w:bdr\s+w:sz="4"\s+w:color="000000"\s*\/>/);
    });

    it('character border passes validator with only size/color', async () => {
      const doc = Document.create();
      const p = new Paragraph();
      const run = new Run('bordered', { border: { size: 6, color: 'FF0000' } });
      p.addRun(run);
      doc.addParagraph(p);
      await expect(doc.toBuffer()).resolves.toBeInstanceOf(Buffer);
      doc.dispose();
    });
  });

  describe('Style-level diagonal cell borders (tl2br / tr2bl)', () => {
    // Exercises `generateBorderElements` in formatting/Style.ts — the shared
    // helper for style-level tblBorders / tcBorders. Validator round-trip
    // skipped here: the Open XML SDK's StyleTableCellProperties class has
    // its own narrower allowed-children restriction that rejects tcBorders
    // in some style contexts regardless of their contents. We just verify
    // the XML shape directly via the Style's toXML.
    it('emits w:val="nil" on tl2br when only size/color set', () => {
      const zipFixture = async (): Promise<void> => {
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
  <w:style w:type="table" w:styleId="DiagTest"><w:name w:val="DiagTest"/></w:style>
</w:styles>`
        );
        zipHandler.addFile(
          'word/document.xml',
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>`
        );
        return await zipHandler.toBuffer().then(() => {});
      };
      zipFixture(); // suppress unused

      // Build style in-memory and stamp cell diagonal borders.
      const { Style } = require('../../src/formatting/Style');
      const style = Style.create({
        styleId: 'DiagTest',
        name: 'DiagTest',
        type: 'table',
      });
      style.setTableCellFormatting({
        borders: {
          tl2br: { size: 4, color: '00FF00' },
          tr2bl: { size: 4, color: '00FF00' },
        },
      });
      const xml = XMLBuilder.elementToString(style.toXML());
      expect(xml).toMatch(/<w:tl2br[^>]*w:val="nil"[^>]*w:sz="4"/);
      expect(xml).toMatch(/<w:tr2bl[^>]*w:val="nil"[^>]*w:sz="4"/);
      // Neither diagonal may emit without w:val.
      expect(xml).not.toMatch(/<w:tl2br\s+w:sz="4"/);
      expect(xml).not.toMatch(/<w:tr2bl\s+w:sz="4"/);
    });
  });
});
