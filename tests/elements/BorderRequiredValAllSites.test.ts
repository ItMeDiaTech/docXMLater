/**
 * Required `w:val` attribute compliance on CT_Border / CT_PageBorder across
 * all remaining emission sites (iteration-70 audit).
 *
 * Per ECMA-376 Part 1 §17.18.2 CT_Border, `w:val` (ST_Border) is REQUIRED.
 * CT_PageBorder (§17.6.2) extends CT_Border so the same rule applies for
 * page-border children of `<w:pgBorders>`.
 *
 * Emission sites fixed in this iteration — each had the same truthy-gate
 * `if (border.style) attrs['w:val'] = border.style` pattern that silently
 * dropped the required attribute when consumers constructed borders with
 * only size / color / space:
 *
 *   - `Style.generateBorderElements` (shared for style-level tblBorders
 *     and tcBorders)
 *   - `Section.toXML` / pgBorders.buildBorder (direct page borders)
 *   - `Section.toXML` / sectPrChange previous-properties pgBorders
 *   - `Table.toXML` / tblPrChange previous-properties tblBorders
 *
 * All four now default `w:val` to `"nil"` (ECMA-376 "no visible border"
 * sentinel) so round-trip through strict OOXML validation stays clean.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Section } from '../../src/elements/Section';
import { Table } from '../../src/elements/Table';
import { Style } from '../../src/formatting/Style';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('CT_Border w:val required — all remaining sites (iter-70)', () => {
  describe('Section page borders (<w:pgBorders>)', () => {
    it('defaults w:val to "nil" when page border sets only size/color', () => {
      const section = new Section({
        pageBorders: {
          offsetFrom: 'page',
          top: { size: 4, color: '000000' },
        },
      });
      const xml = XMLBuilder.elementToString(section.toXML());
      expect(xml).toMatch(/<w:pgBorders[^>]*><w:top[^>]*w:val="nil"[^>]*w:sz="4"/);
      expect(xml).not.toMatch(/<w:top\s+w:sz="4"\s+w:color="000000"\s*\/>/);
    });

    it('page border passes OOXML validator with only size/color', async () => {
      // Load a doc that carries a page border set, then (via the generator
      // roundtrip) re-emit. Since the default SectionProperties from the
      // XML fixture already has explicit `style="single"`, we build the
      // fixture with a border missing `w:val` on the generator side by
      // rewriting the Section's properties via load-time parse + direct
      // toXML re-emission through Document.toBuffer().
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="4" w:color="000000"/></w:pgBorders></w:sectPr></w:body></w:document>`
      );
      const buffer = await zipHandler.toBuffer();

      const doc = await Document.loadFromBuffer(buffer);
      // Access the section and overwrite pageBorders via the section's
      // private properties — matches what a programmatic style-update
      // pipeline would do. Cast through `any` since there is no public
      // setter for pageBorders.
      const section = doc.getSection();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (section as any).properties.pageBorders = {
        offsetFrom: 'page',
        top: { size: 4, color: 'FF0000' },
        left: { size: 4, color: 'FF0000' },
      };
      await expect(doc.toBuffer()).resolves.toBeInstanceOf(Buffer);
      doc.dispose();
    });
  });

  describe('Style-level tblBorders / tcBorders (<w:tblPr>/<w:tcPr>)', () => {
    it('style tblBorders passes validator with borders missing style', async () => {
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
  <w:style w:type="table" w:styleId="TBorder"><w:name w:val="TBorder"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`
      );
      zipHandler.addFile(
        'word/document.xml',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>`
      );
      const buffer = await zipHandler.toBuffer();

      const doc = await Document.loadFromBuffer(buffer);
      const style = doc.getStylesManager().getStyle('TBorder') as Style;
      const existing = style.getProperties().tableStyle ?? {};
      // Mutate: strip the style field so the regenerator encounters a
      // size/color-only border and must supply a default w:val.
      style.setTableFormatting({
        ...(existing.table ?? {}),
        borders: { top: { size: 4, color: 'FF0000' } },
      });
      await expect(doc.toBuffer()).resolves.toBeInstanceOf(Buffer);
      doc.dispose();
    });
  });

  describe('Table tblPrChange previous borders', () => {
    it('tblPrChange previous tblBorders passes validator with borders missing style', async () => {
      const doc = Document.create();
      const table = new Table(2, 2);
      // Set current borders to something.
      table.setBorders({ top: { style: 'single', size: 4, color: '000000' } });
      // Set a previous-borders snapshot that lacks `style` on one side.
      table.setTblPrChange({
        id: '1',
        author: 'Tester',
        date: '2026-01-01T00:00:00Z',
        previousProperties: {
          borders: { top: { size: 4, color: 'FF0000' } },
        },
      });
      doc.addTable(table);
      await expect(doc.toBuffer()).resolves.toBeInstanceOf(Buffer);
      doc.dispose();
    });
  });
});
