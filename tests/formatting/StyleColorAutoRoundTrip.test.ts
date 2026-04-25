/**
 * Style-level run color `<w:color w:val="auto"/>` round-trip.
 *
 * Per ECMA-376 Part 1 §17.3.2.6 CT_Color, `w:val` is ST_HexColor which
 * allows either a 6-hex RGB value or the special sentinel `"auto"`
 * (ST_HexColorAuto). `"auto"` tells Word to render using the automatic /
 * window text color (typically black, but adapts to dark mode and
 * high-contrast themes).
 *
 * The style-level rPr parser (`parseRunFormattingFromXml`) previously
 * dropped `w:val="auto"` with an explicit `val !== 'auto'` exclusion,
 * presumably to avoid `normalizeColor()` rejecting non-hex input. As a
 * result, any paragraph/character style carrying `<w:color w:val="auto"/>`
 * silently lost that marker on load; the emitter then defaulted to
 * `"000000"` — forcing the style to render as pure black instead of
 * "inherit auto color", which changes visual output in dark-mode /
 * high-contrast environments.
 *
 * The direct-run rPr parser (`parseRunFromObject`) already handled this
 * correctly (§17.18.6 path). This test exercises the style-level path to
 * confirm the fix preserves the literal "auto" through round-trip.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyleColor(colorVal: string): Promise<Buffer> {
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
  <w:style w:type="character" w:styleId="AutoColorStyle">
    <w:name w:val="AutoColorStyle"/>
    <w:rPr><w:color w:val="${colorVal}"/></w:rPr>
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

describe('Style-level <w:color w:val="auto"/> round-trip (§17.3.2.6 / §17.18.6)', () => {
  it('parses w:val="auto" into style properties', async () => {
    const buffer = await makeDocxWithStyleColor('auto');
    const doc = await Document.loadFromBuffer(buffer);
    const style = doc.getStylesManager().getStyle('AutoColorStyle');
    expect(style?.getProperties().runFormatting?.color).toBe('auto');
    doc.dispose();
  });

  it('round-trips w:val="auto" through save — emits <w:color w:val="auto"/> not "000000"', async () => {
    const buffer = await makeDocxWithStyleColor('auto');
    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const stylesFile = zip.getFile('word/styles.xml');
    const content = stylesFile?.content;
    const styles = content instanceof Buffer ? content.toString('utf8') : String(content);

    // Re-emitted style rPr must carry `w:val="auto"` — previously the
    // parser dropped it and the emitter defaulted to "000000", changing
    // the rendering semantic from "use automatic color" to "always black".
    expect(styles).toMatch(/<w:rPr>[\s\S]*<w:color\s+w:val="auto"/);
    expect(styles).not.toMatch(/<w:rPr>[\s\S]*<w:color\s+w:val="000000"/);
  });

  it('still parses concrete hex values correctly (regression guard)', async () => {
    const buffer = await makeDocxWithStyleColor('FF0000');
    const doc = await Document.loadFromBuffer(buffer);
    const style = doc.getStylesManager().getStyle('AutoColorStyle');
    expect(style?.getProperties().runFormatting?.color).toBe('FF0000');
    doc.dispose();
  });
});
