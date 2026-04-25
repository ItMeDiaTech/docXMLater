/**
 * `<w:hyperlink w:history="…"/>` — ST_OnOff round-trip.
 *
 * Per ECMA-376 Part 1 §17.16.22 CT_Hyperlink, `w:history` is a CT_OnOff
 * attribute controlling whether the hyperlink is added to the browser
 * history list. It accepts every ST_OnOff literal
 * ("1"/"0"/"true"/"false"/"on"/"off").
 *
 * Two bugs compounded:
 *   1. XMLParser's `parseAttributeValue: true` coerces `"0"` to the
 *      number `0` and `"false"` to the boolean `false`.
 *   2. The Hyperlink emitter wrote the raw stored value through a
 *      `if (this.history)` truthy check, so the coerced falsy forms
 *      were silently dropped from the output.
 *
 * Net effect: `w:history="0"` (or `w:history="false"` / `w:history="off"`)
 * was lost on every load → save round-trip — the author's explicit
 * "don't add to history" override silently reverted to the spec
 * default of "absent" (i.e., add to history).
 *
 * Iteration 100 normalises the parsed value to "1"/"0" via
 * `parseOnOffAttribute` and changes the emitter guard to `!== undefined`.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadAndResaveDocXml(xml: string): Promise<string> {
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
  zipHandler.addFile('word/document.xml', xml);
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const out = await doc.toBuffer();
  doc.dispose();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(out);
  const content = zip.getFile('word/document.xml')?.content;
  return content instanceof Buffer ? content.toString('utf8') : String(content);
}

function buildHyperlinkDoc(historyAttr: string | null): string {
  const attrFragment = historyAttr === null ? '' : ` w:history="${historyAttr}"`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink w:anchor="bookmark1"${attrFragment}>
        <w:r><w:t>link</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;
}

function extractHyperlinkOpenTag(xml: string): string {
  return xml.match(/<w:hyperlink[^>]*>/)?.[0] ?? '';
}

describe('<w:hyperlink w:history> ST_OnOff round-trip', () => {
  it('preserves w:history="1" (baseline)', async () => {
    const out = await loadAndResaveDocXml(buildHyperlinkDoc('1'));
    expect(extractHyperlinkOpenTag(out)).toMatch(/w:history="1"/);
  });

  it('preserves w:history="0" — previously silently dropped', async () => {
    const out = await loadAndResaveDocXml(buildHyperlinkDoc('0'));
    const tag = extractHyperlinkOpenTag(out);
    expect(tag).toMatch(/w:history="0"/);
  });

  it('normalises w:history="true" to "1"', async () => {
    const out = await loadAndResaveDocXml(buildHyperlinkDoc('true'));
    expect(extractHyperlinkOpenTag(out)).toMatch(/w:history="1"/);
  });

  it('normalises w:history="false" to "0" — previously silently dropped', async () => {
    const out = await loadAndResaveDocXml(buildHyperlinkDoc('false'));
    const tag = extractHyperlinkOpenTag(out);
    expect(tag).toMatch(/w:history="0"/);
  });

  it('normalises w:history="on" to "1"', async () => {
    const out = await loadAndResaveDocXml(buildHyperlinkDoc('on'));
    expect(extractHyperlinkOpenTag(out)).toMatch(/w:history="1"/);
  });

  it('normalises w:history="off" to "0"', async () => {
    const out = await loadAndResaveDocXml(buildHyperlinkDoc('off'));
    const tag = extractHyperlinkOpenTag(out);
    expect(tag).toMatch(/w:history="0"/);
  });

  it('omits w:history attribute entirely when absent (regression guard)', async () => {
    const out = await loadAndResaveDocXml(buildHyperlinkDoc(null));
    const tag = extractHyperlinkOpenTag(out);
    expect(tag).not.toMatch(/w:history=/);
  });
});
