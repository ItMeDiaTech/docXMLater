/**
 * Floating table positioning `<w:tblpPr>` — zero-value attributes
 * must round-trip (parser/emitter symmetry).
 *
 * Per ECMA-376 Part 1 §17.4.52 CT_TblPPr, six numeric attributes can
 * legitimately be zero:
 *   - `w:tblpX`           ST_SignedTwipsMeasure (x offset; 0 = at anchor)
 *   - `w:tblpY`           ST_SignedTwipsMeasure (y offset; 0 = at anchor)
 *   - `w:leftFromText`    ST_TwipsMeasure       (text-distance, 0 valid)
 *   - `w:rightFromText`   ST_TwipsMeasure
 *   - `w:topFromText`     ST_TwipsMeasure
 *   - `w:bottomFromText`  ST_TwipsMeasure
 *
 * Bug: both the main-path parser (DocumentParser.ts §6836) and the
 * `parseGenericPreviousProperties` variant (§11440) gated each read
 * with a truthy check — `if (tblpPr['@_w:tblpX']) position.x = …`.
 * XMLParser coerces `"0"` to the number `0`, which is falsy, so every
 * zero-valued attribute was silently dropped on load.
 *
 * The emitters (Table.ts §1653) already use `!== undefined`, so the
 * parser/emitter asymmetry lost zero-offset floating tables on
 * round-trip and dropped tracked-change "previous" zero positions
 * inside `w:tblPrChange`.
 *
 * Iteration 132 routes each numeric tblpPr attribute through
 * `isExplicitlySet` + `safeParseInt` on BOTH parsers so zero values
 * survive.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadDocWith(tblPrInnerXml: string) {
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
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        ${tblPrInnerXml}
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>c</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>d</w:t></w:r></w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  return Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
}

describe('<w:tblpPr> zero-value attribute round-trip', () => {
  it('preserves w:tblpX="0" (floating table anchored at x=0)', async () => {
    const doc = await loadDocWith(
      '<w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="0" w:tblpY="100"/>'
    );
    const table = doc.getTables()[0]!;
    const pos = table.getPosition();
    doc.dispose();
    expect(pos).toBeDefined();
    expect(pos!.x).toBe(0);
    expect(pos!.y).toBe(100);
  });

  it('preserves w:tblpY="0"', async () => {
    const doc = await loadDocWith(
      '<w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="100" w:tblpY="0"/>'
    );
    const table = doc.getTables()[0]!;
    const pos = table.getPosition();
    doc.dispose();
    expect(pos!.y).toBe(0);
    expect(pos!.x).toBe(100);
  });

  it('preserves w:leftFromText="0" and w:rightFromText="0"', async () => {
    const doc = await loadDocWith(
      '<w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="200" w:tblpY="300" w:leftFromText="0" w:rightFromText="0"/>'
    );
    const table = doc.getTables()[0]!;
    const pos = table.getPosition();
    doc.dispose();
    expect(pos!.leftFromText).toBe(0);
    expect(pos!.rightFromText).toBe(0);
  });

  it('preserves w:topFromText="0" and w:bottomFromText="0"', async () => {
    const doc = await loadDocWith(
      '<w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="10" w:tblpY="20" w:topFromText="0" w:bottomFromText="0"/>'
    );
    const table = doc.getTables()[0]!;
    const pos = table.getPosition();
    doc.dispose();
    expect(pos!.topFromText).toBe(0);
    expect(pos!.bottomFromText).toBe(0);
  });

  it('preserves all six numeric attributes at zero (edge case)', async () => {
    const doc = await loadDocWith(
      '<w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="0" w:tblpY="0" w:leftFromText="0" w:rightFromText="0" w:topFromText="0" w:bottomFromText="0"/>'
    );
    const table = doc.getTables()[0]!;
    const pos = table.getPosition();
    doc.dispose();
    expect(pos).toBeDefined();
    expect(pos!.x).toBe(0);
    expect(pos!.y).toBe(0);
    expect(pos!.leftFromText).toBe(0);
    expect(pos!.rightFromText).toBe(0);
    expect(pos!.topFromText).toBe(0);
    expect(pos!.bottomFromText).toBe(0);
  });

  it('preserves zero w:tblpX in tblPrChange previous state', async () => {
    const doc = await loadDocWith(
      `<w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="500" w:tblpY="500"/>
       <w:tblPrChange w:id="1" w:author="A" w:date="2026-04-24T10:00:00Z">
         <w:tblPr>
           <w:tblW w:w="5000" w:type="pct"/>
           <w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="0" w:tblpY="0" w:leftFromText="0"/>
         </w:tblPr>
       </w:tblPrChange>`
    );
    const table = doc.getTables()[0]!;
    const tblPrChange = table.getTblPrChange();
    doc.dispose();
    expect(tblPrChange).toBeDefined();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const prevPos = (tblPrChange!.previousProperties as any).position;
    expect(prevPos).toBeDefined();
    expect(prevPos.x).toBe(0);
    expect(prevPos.y).toBe(0);
    expect(prevPos.leftFromText).toBe(0);
  });

  it('preserves positive values (regression guard)', async () => {
    const doc = await loadDocWith(
      '<w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="1440" w:tblpY="2880" w:leftFromText="180" w:rightFromText="180"/>'
    );
    const table = doc.getTables()[0]!;
    const pos = table.getPosition();
    doc.dispose();
    expect(pos!.x).toBe(1440);
    expect(pos!.y).toBe(2880);
    expect(pos!.leftFromText).toBe(180);
    expect(pos!.rightFromText).toBe(180);
  });
});
