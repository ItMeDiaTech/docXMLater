/**
 * Style-level `<w:pBdr>` — parser must capture all CT_Border attributes
 * per ECMA-376 Part 1 §17.18.2 (nine attrs: val, sz, space, color,
 * themeColor, themeTint, themeShade, shadow, frame).
 *
 * The style **emitter** (`Style.generateParagraphProperties`) already
 * emits all nine. The style **parser**
 * (`DocumentParser.parseParagraphFormattingFromXml` — the XML-string
 * variant used for styles.xml) read only four (style, size, space,
 * color), silently dropping themeColor / themeTint / themeShade /
 * shadow / frame on round-trip.
 *
 * Iteration 98 extends the style-level parser to read all nine, with
 * shadow/frame routed through `parseOnOffAttribute` so ST_OnOff
 * literals (`"on"` / `"off"` / `"true"` / `"false"`) resolve.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadStyleBordersFromStylesXml(
  stylesXmlBody: string
): Promise<Record<string, any> | undefined> {
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
  ${stylesXmlBody}
</w:styles>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const style = doc.getStylesManager().getStyle('Bordered');
  const pf = style?.getParagraphFormatting();
  doc.dispose();
  return pf?.borders as Record<string, any> | undefined;
}

describe('style-level <w:pBdr> CT_Border full attribute coverage', () => {
  it('parses themeColor / themeTint / themeShade on style-level paragraph borders', async () => {
    const body = `
      <w:style w:type="paragraph" w:styleId="Bordered">
        <w:name w:val="Bordered"/>
        <w:pPr>
          <w:pBdr>
            <w:top w:val="single" w:sz="4" w:space="1" w:color="auto" w:themeColor="accent1" w:themeTint="80" w:themeShade="40"/>
            <w:bottom w:val="double" w:sz="6" w:space="2" w:color="FF0000"/>
          </w:pBdr>
        </w:pPr>
      </w:style>`;
    const borders = await loadStyleBordersFromStylesXml(body);
    expect(borders?.top?.themeColor).toBe('accent1');
    expect(borders?.top?.themeTint).toBe('80');
    expect(borders?.top?.themeShade).toBe('40');
    expect(borders?.top?.style).toBe('single');
    expect(borders?.bottom?.style).toBe('double');
    expect(borders?.bottom?.color).toBe('FF0000');
  });

  it('parses shadow/frame booleans on style-level paragraph borders', async () => {
    const body = `
      <w:style w:type="paragraph" w:styleId="Bordered">
        <w:name w:val="Bordered"/>
        <w:pPr>
          <w:pBdr>
            <w:top w:val="single" w:sz="8" w:color="000000" w:shadow="1" w:frame="0"/>
          </w:pBdr>
        </w:pPr>
      </w:style>`;
    const borders = await loadStyleBordersFromStylesXml(body);
    expect(borders?.top?.shadow).toBe(true);
    expect(borders?.top?.frame).toBe(false);
  });

  it('honours ST_OnOff "on"/"off" literals on style-level shadow/frame', async () => {
    const body = `
      <w:style w:type="paragraph" w:styleId="Bordered">
        <w:name w:val="Bordered"/>
        <w:pPr>
          <w:pBdr>
            <w:top w:val="single" w:sz="4" w:color="auto" w:shadow="on" w:frame="off"/>
          </w:pBdr>
        </w:pPr>
      </w:style>`;
    const borders = await loadStyleBordersFromStylesXml(body);
    expect(borders?.top?.shadow).toBe(true);
    expect(borders?.top?.frame).toBe(false);
  });
});
