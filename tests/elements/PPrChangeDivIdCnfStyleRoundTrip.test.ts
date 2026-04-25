/**
 * `<w:pPrChange><w:pPr><w:divId …/> + <w:cnfStyle …/>` — tracked-change
 * history of paragraph divId and cnfStyle must round-trip.
 *
 * Per ECMA-376 Part 1:
 *   - §17.3.1.10 CT_DivId: `w:val` is ST_DecimalNumber (xsd:integer);
 *     0 is a legal ID referencing the first `<w:div>` in web settings.
 *   - §17.3.1.8 CT_Cnf: `w:val` is ST_Cnf (12-character bitmask)
 *     identifying which conditional-formatting flags from the parent
 *     table style apply to the paragraph.
 *
 * Both fields are preserved by the pPrChange emitter at
 * `Paragraph.ts:3915-3921`. The parser at `DocumentParser.ts:2880`
 * never read either, so a tracked change recording a previous divId
 * or cnfStyle silently lost it on load → save — breaking Word's
 * "Original" markup view for documents that track these properties.
 *
 * Iteration 115 mirrors the main-pPr parse semantics onto the
 * pPrChange path: `isExplicitlySet` + `safeParseInt` for divId (so
 * id=0 survives XMLParser numeric coercion), and
 * `String().padStart(12, '0')` for cnfStyle (defensive normalisation
 * of any numerically-coerced short forms).
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

function extractPPrChangeBlock(xml: string): string {
  return xml.match(/<w:pPrChange[\s\S]*?<\/w:pPrChange>/)?.[0] ?? '';
}

describe('<w:pPrChange> previous <w:divId> / <w:cnfStyle> round-trip', () => {
  it('preserves previous divId (positive value)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:pPrChange w:id="1" w:author="T" w:date="2026-01-01T00:00:00Z">
          <w:pPr>
            <w:divId w:val="42"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>content</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const block = extractPPrChangeBlock(out);
    expect(block).toMatch(/<w:divId[^/]*w:val="42"/);
  });

  it('parses previous divId="0" on load (zero is a legal ID per §17.3.1.10)', async () => {
    // Parser-side assertion only: the SDK validator rejects divId
    // references to divs not declared in webSettings.xml (orthogonal
    // schema-reference concern), so we verify the in-memory state
    // rather than round-tripping through toBuffer. This mirrors the
    // iter-106 main-pPr divId=0 test approach.
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
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:pPrChange w:id="2" w:author="T" w:date="2026-01-01T00:00:00Z">
          <w:pPr>
            <w:divId w:val="0"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>content</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
    );
    const buffer = await zipHandler.toBuffer();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const paragraph = doc.getParagraphs()[0]!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const change = (paragraph as any).formatting.pPrChange as
      | { previousProperties?: { divId?: number } }
      | undefined;
    doc.dispose();
    expect(change?.previousProperties?.divId).toBe(0);
  });

  it('preserves previous cnfStyle 12-char bitmask', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="TableGrid"/>
        <w:pPrChange w:id="3" w:author="T" w:date="2026-01-01T00:00:00Z">
          <w:pPr>
            <w:cnfStyle w:val="100000000000"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>content</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const block = extractPPrChangeBlock(out);
    expect(block).toMatch(/<w:cnfStyle[^/]*w:val="100000000000"/);
  });
});
