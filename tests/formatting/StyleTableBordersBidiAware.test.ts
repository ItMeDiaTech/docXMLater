/**
 * Style-level `<w:tblBorders>` / `<w:tcBorders>` — bidi-aware
 * `<w:start>` / `<w:end>` border aliases must parse.
 *
 * Per ECMA-376 Part 1 §17.4.40 CT_TblBorders and §17.4.66 CT_TcBorders,
 * the left/right border elements each have a bidi-aware alias:
 *   - `<w:start>` is the preferred spelling; `<w:left>` is the legacy
 *     LTR-only form.
 *   - `<w:end>` is the preferred spelling; `<w:right>` is the legacy
 *     LTR-only form.
 *
 * Modern authoring tools (Word 2013+, Google Docs) emit the bidi-aware
 * form by default. Previously `parseBordersFromXml` only matched
 * `w:left` / `w:right`, so any table or cell style authored with
 * `<w:start>` / `<w:end>` in its borders silently lost those sides on
 * load. The style emitter always writes `w:left` / `w:right` (the
 * internal model keys), so on re-save the bidi-aware forms were
 * replaced wholesale by blank LTR output — leaving the authored
 * table with no side borders at all.
 *
 * Iteration 121 extends the parser to prefer `w:start`/`w:end` when
 * present, storing under the `left`/`right` keys — matching the
 * existing CT_Ind and CT_TblCellMar bidi-aware fallback pattern.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadStyleBorders(stylesXmlBody: string): Promise<Record<string, any>> {
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
  const style = doc.getStylesManager().getStyle('TableStyle');
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const tblStyle = (style as any)?.properties?.tableStyle;
  doc.dispose();
  return tblStyle?.table?.borders ?? {};
}

describe('table style borders bidi-aware <w:start>/<w:end> parsing', () => {
  it('parses <w:start>/<w:end> into left/right keys (preferred form)', async () => {
    const body = `
      <w:style w:type="table" w:styleId="TableStyle">
        <w:name w:val="TableStyle"/>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:start w:val="double" w:sz="8" w:color="FF0000"/>
            <w:bottom w:val="single" w:sz="4" w:color="000000"/>
            <w:end w:val="double" w:sz="8" w:color="00FF00"/>
          </w:tblBorders>
        </w:tblPr>
      </w:style>`;
    const borders = await loadStyleBorders(body);
    expect(borders.left?.style).toBe('double');
    expect(borders.left?.color).toBe('FF0000');
    expect(borders.right?.style).toBe('double');
    expect(borders.right?.color).toBe('00FF00');
  });

  it('parses legacy <w:left>/<w:right> (regression guard)', async () => {
    const body = `
      <w:style w:type="table" w:styleId="TableStyle">
        <w:name w:val="TableStyle"/>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:left w:val="dashed" w:sz="6" w:color="AA0000"/>
            <w:bottom w:val="single" w:sz="4" w:color="000000"/>
            <w:right w:val="dashed" w:sz="6" w:color="00AA00"/>
          </w:tblBorders>
        </w:tblPr>
      </w:style>`;
    const borders = await loadStyleBorders(body);
    expect(borders.left?.style).toBe('dashed');
    expect(borders.left?.color).toBe('AA0000');
    expect(borders.right?.style).toBe('dashed');
    expect(borders.right?.color).toBe('00AA00');
  });

  it('prefers <w:start> over <w:left> when both appear (spec precedence)', async () => {
    // Per §17.4.40 the bidi-aware spelling takes precedence.
    const body = `
      <w:style w:type="table" w:styleId="TableStyle">
        <w:name w:val="TableStyle"/>
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:start w:val="double" w:sz="8" w:color="FF0000"/>
            <w:left w:val="dotted" w:sz="2" w:color="999999"/>
            <w:bottom w:val="single" w:sz="4" w:color="000000"/>
          </w:tblBorders>
        </w:tblPr>
      </w:style>`;
    const borders = await loadStyleBorders(body);
    // The bidi-aware "start" should win.
    expect(borders.left?.style).toBe('double');
    expect(borders.left?.color).toBe('FF0000');
  });
});
