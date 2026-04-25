/**
 * `<w:sectPrChange>` previous CT_OnOff flags — explicit-false values
 * must round-trip alongside explicit-true.
 *
 * Per ECMA-376 Part 1, the following section-level flags are CT_OnOff
 * (§17.17.4 — accept every ST_OnOff literal "1"/"0"/"true"/"false"/
 * "on"/"off"):
 *   - w:formProt (§17.6.9) — form protection
 *   - w:noEndnote (§17.11.14) — suppress endnotes in section
 *   - w:titlePg (§17.10.6) — distinct first-page header/footer
 *   - w:bidi (§17.6.1) — right-to-left section layout
 *   - w:rtlGutter (§17.6.16) — book-binding gutter on the right
 *
 * The main-sectPr emitter at `Section.ts:1286-1313` already preserved
 * the explicit-false distinction via `!== undefined`. The
 * sectPrChange emitter at `Section.ts:1566-1600`, however, used a
 * truthy gate on each flag:
 *
 *     if (prev.formProt) { …emit… }
 *
 * so a tracked change capturing "previous state was FALSE" (e.g.,
 * form protection was OFF before the revision that turned it ON)
 * silently dropped the `<w:formProt w:val="0"/>` marker on save. The
 * parser already produces `false` for explicit-false input via
 * `parseSectCtOnOff`, so the bug manifested as an emitter-side drop
 * only — breaking Word's "Original" view for the previous state.
 *
 * Iteration 117 brings the sectPrChange emitter to parity with the
 * main-sectPr emitter for all five flags.
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
  const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
  const out = await doc.toBuffer();
  doc.dispose();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(out);
  const content = zip.getFile('word/document.xml')?.content;
  return content instanceof Buffer ? content.toString('utf8') : String(content);
}

function extractSectPrChange(xml: string): string {
  return xml.match(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/)?.[0] ?? '';
}

describe('<w:sectPrChange> previous CT_OnOff explicit-false round-trip', () => {
  it('preserves previous w:formProt w:val="0" (explicit false)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:formProt w:val="1"/>
      <w:sectPrChange w:id="1" w:author="T" w:date="2026-01-01T00:00:00Z">
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          <w:formProt w:val="0"/>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = extractSectPrChange(out);
    // Previously the explicit-false formProt was silently dropped.
    expect(changeBlock).toMatch(/<w:formProt[^/]*w:val="0"/);
  });

  it('preserves previous w:titlePg w:val="0" (explicit false)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:titlePg w:val="1"/>
      <w:sectPrChange w:id="2" w:author="T" w:date="2026-01-01T00:00:00Z">
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          <w:titlePg w:val="0"/>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = extractSectPrChange(out);
    expect(changeBlock).toMatch(/<w:titlePg[^/]*w:val="0"/);
  });

  it('preserves previous w:bidi w:val="0" (explicit false LTR override)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:bidi w:val="1"/>
      <w:sectPrChange w:id="3" w:author="T" w:date="2026-01-01T00:00:00Z">
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          <w:bidi w:val="0"/>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = extractSectPrChange(out);
    expect(changeBlock).toMatch(/<w:bidi[^/]*w:val="0"/);
  });

  it('preserves explicit-true (regression guard)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:sectPrChange w:id="4" w:author="T" w:date="2026-01-01T00:00:00Z">
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          <w:formProt w:val="1"/>
          <w:titlePg/>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = extractSectPrChange(out);
    // formProt: bare (implicit true) form preserved
    expect(changeBlock).toMatch(/<w:formProt(\/>|[^/]*w:val="1"[^/]*\/>)/);
    // titlePg: explicit true with w:val="1"
    expect(changeBlock).toMatch(/<w:titlePg[^/]*w:val="1"/);
  });
});
