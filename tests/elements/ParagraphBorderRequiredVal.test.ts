/**
 * Paragraph border `<w:top>`/`<w:left>`/`<w:bottom>`/`<w:right>`/`<w:between>`/
 * `<w:bar>` inside `<w:pBdr>` — required `w:val` attribute compliance.
 *
 * Per ECMA-376 Part 1 §17.18.2 CT_Border, `w:val` (ST_Border) is REQUIRED.
 * Strict OOXML validation rejects any border emission missing it with
 * *"The required attribute 'val' is missing"*.
 *
 * Three separate emission paths had the same bug:
 *   1. `Paragraph.generateParagraphPropertiesXml` — direct paragraph borders
 *   2. `Paragraph.generateParagraphPropertiesXml` — pPrChange/previousProperties borders
 *   3. `Style.generateParagraphProperties` — style-level paragraph borders
 *
 * Each used a local truthy-check `if (border.style)` that silently dropped
 * `w:val` whenever a consumer constructed a border with only size/color (the
 * fields most people think about) and no explicit style. Fix: default
 * `w:val` to `"nil"` (the ECMA-376 ST_Border "no border" value) so the
 * output is always schema-compliant. Matches the behavior of
 * `XMLBuilder.createBorder` used by TableCell / Table which defaults to
 * `"single"`, and `TableRow.buildBordersXML` (fixed in iteration 68) which
 * uses `"nil"`.
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('Paragraph pBdr <w:top>/etc. — required w:val compliance (§17.18.2)', () => {
  it('defaults w:val to "nil" when direct paragraph border has no style', () => {
    const p = new Paragraph();
    p.addText('x');
    p.setBorder({ top: { size: 4, color: '000000' } });
    const xml = XMLBuilder.elementToString(p.toXML());
    // Direct paragraph pBdr/top must carry w:val even when consumer only set size+color.
    expect(xml).toMatch(/<w:pBdr>[\s\S]*?<w:top[^>]*w:val="nil"[^>]*w:sz="4"[\s\S]*?<\/w:pBdr>/);
    // Must NOT emit a top border missing w:val.
    expect(xml).not.toMatch(/<w:top\s+w:sz="4"\s+w:color="000000"\s*\/>/);
  });

  it('direct paragraph border passes OOXML validator with only size/color', async () => {
    // Construct via Document path so toBuffer() runs validation.
    const doc = Document.create();
    const p = new Paragraph();
    p.addText('x');
    p.setBorder({ top: { size: 4, color: '000000' } });
    doc.addParagraph(p);
    await expect(doc.toBuffer()).resolves.toBeInstanceOf(Buffer);
    doc.dispose();
  });

  it('style-level paragraph border defaults w:val to "nil" when style omitted', async () => {
    // Direct test of the style-level paragraph pBdr generator path. Construct
    // a docx whose styles.xml defines a pPr with a top border missing w:val,
    // then round-trip through load → toBuffer so the generator path re-emits.
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
  <w:style w:type="paragraph" w:styleId="BorderedPara">
    <w:name w:val="BorderedPara"/>
    <w:pPr>
      <w:pBdr>
        <w:top w:val="single" w:sz="4" w:color="000000"/>
      </w:pBdr>
    </w:pPr>
  </w:style>
</w:styles>`
    );
    zipHandler.addFile(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body>
</w:document>`
    );
    const buffer = await zipHandler.toBuffer();

    const doc = await Document.loadFromBuffer(buffer);
    // Mutate the style's paragraph formatting to carry a border with no style
    // (size/color only). The serializer must default `w:val="nil"` or the
    // validator rejects the output.
    const style = doc.getStylesManager().getStyle('BorderedPara')!;
    const existing = style.getProperties().paragraphFormatting ?? {};
    style.setParagraphFormatting({
      ...existing,
      borders: { top: { size: 4, color: 'FF0000' } },
    });
    await expect(doc.toBuffer()).resolves.toBeInstanceOf(Buffer);
    doc.dispose();
  });
});
