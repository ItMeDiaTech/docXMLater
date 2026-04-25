/**
 * Charts (`c:chartSpace`), SmartArt (`a:graphic` with diagram data), and
 * OLE objects (`w:object`) are stored in a DOCX as a combination of:
 *   - The reference inside `word/document.xml` (a `w:drawing` /
 *     `w:object` element with a relationship-id pointing to the part)
 *   - The part itself: `word/charts/chart*.xml`, `word/diagrams/data*.xml`
 *     + `word/diagrams/layout*.xml` + `word/diagrams/quickStyle*.xml` +
 *     `word/diagrams/colors*.xml`, or `word/embeddings/oleObject*.bin`
 *   - A relationship from `word/_rels/document.xml.rels` to that part
 *   - Content_Types overrides
 *
 * docxmlater has no editing API for these but must preserve them
 * verbatim on round-trip — the framework's promise is that documents
 * survive load → save unchanged.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

const CHART_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
  `<c:chart>` +
  `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Sales Q1</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` +
  `<c:plotArea>` +
  `<c:layout/>` +
  `<c:barChart>` +
  `<c:barDir val="bar"/>` +
  `<c:grouping val="clustered"/>` +
  `<c:varyColors val="0"/>` +
  `<c:ser>` +
  `<c:idx val="0"/>` +
  `<c:order val="0"/>` +
  `<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Series 1</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
  `<c:val><c:numRef><c:f>Sheet1!$B$2:$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val>` +
  `</c:ser>` +
  `<c:axId val="1"/>` +
  `<c:axId val="2"/>` +
  `</c:barChart>` +
  `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="2"/></c:catAx>` +
  `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="1"/></c:valAx>` +
  `</c:plotArea>` +
  `<c:plotVisOnly val="1"/>` +
  `<c:dispBlanksAs val="gap"/>` +
  `</c:chart>` +
  `</c:chartSpace>`;

const DIAGRAM_DATA_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">` +
  `<dgm:ptLst><dgm:pt modelId="{00000000-0000-0000-0000-000000000001}" type="doc"/></dgm:ptLst>` +
  `</dgm:dataModel>`;

const OLE_BIN = Buffer.from('OLE_BINARY_PAYLOAD_PLACEHOLDER', 'utf8');

async function buildDocxWithEmbedded(): Promise<Buffer> {
  // Use Document.create() as a valid OOXML base, then add chart/diagram/OLE
  // parts at the ZIP layer. Round-trip must keep them all.
  const seed = Document.create();
  seed.createParagraph('Embedded test');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);

  zip.addFile('word/charts/chart1.xml', CHART_XML);
  zip.addFile('word/diagrams/data1.xml', DIAGRAM_DATA_XML);
  zip.addFile('word/embeddings/oleObject1.bin', OLE_BIN, { binary: true });

  // Wire Content_Types overrides.
  const ct = zip.getFileAsString('[Content_Types].xml')!;
  const updatedCt = ct.replace(
    '</Types>',
    `<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>` +
      `<Override PartName="/word/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>` +
      `<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>` +
      `</Types>`
  );
  zip.updateFile('[Content_Types].xml', updatedCt);

  // Wire document → chart / diagram / oleObject relationships.
  const docRels = zip.getFileAsString('word/_rels/document.xml.rels')!;
  const updatedRels = docRels.replace(
    '</Relationships>',
    `<Relationship Id="rId500" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>` +
      `<Relationship Id="rId501" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>` +
      `<Relationship Id="rId502" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>` +
      `</Relationships>`
  );
  zip.updateFile('word/_rels/document.xml.rels', updatedRels);

  return zip.toBuffer();
}

describe('Chart / SmartArt / OLE passthrough', () => {
  it('preserves chart, diagram, and OLE parts on unmodified round-trip', async () => {
    const buf1 = await buildDocxWithEmbedded();
    const doc = await Document.loadFromBuffer(buf1);
    const buf2 = await doc.toBuffer();
    doc.dispose();

    const out = new ZipHandler();
    await out.loadFromBuffer(buf2);
    expect(out.hasFile('word/charts/chart1.xml')).toBe(true);
    expect(out.hasFile('word/diagrams/data1.xml')).toBe(true);
    expect(out.hasFile('word/embeddings/oleObject1.bin')).toBe(true);
    expect(out.getFileAsString('word/charts/chart1.xml')).toContain('Sales Q1');
    expect(out.getFileAsString('word/diagrams/data1.xml')).toContain(
      '{00000000-0000-0000-0000-000000000001}'
    );
  });

  it('preserves the chart/diagram/oleObject relationships on the document part', async () => {
    const buf1 = await buildDocxWithEmbedded();
    const doc = await Document.loadFromBuffer(buf1);
    const buf2 = await doc.toBuffer();
    doc.dispose();

    const out = new ZipHandler();
    await out.loadFromBuffer(buf2);
    const rels = out.getFileAsString('word/_rels/document.xml.rels')!;
    expect(rels).toMatch(/Type="[^"]*relationships\/chart"/);
    expect(rels).toMatch(/Type="[^"]*relationships\/diagramData"/);
    expect(rels).toMatch(/Type="[^"]*relationships\/oleObject"/);
  });

  it('preserves Content_Types entries for chart/diagram parts', async () => {
    const buf1 = await buildDocxWithEmbedded();
    const doc = await Document.loadFromBuffer(buf1);
    const buf2 = await doc.toBuffer();
    doc.dispose();

    const out = new ZipHandler();
    await out.loadFromBuffer(buf2);
    const ct = out.getFileAsString('[Content_Types].xml')!;
    expect(ct).toContain('drawingml.chart+xml');
    expect(ct).toContain('drawingml.diagramData+xml');
  });

  it('preserves all parts when document body is modified between load and save', async () => {
    const buf1 = await buildDocxWithEmbedded();
    const doc = await Document.loadFromBuffer(buf1);
    doc.createParagraph('Modified before save');
    const buf2 = await doc.toBuffer();
    doc.dispose();

    const out = new ZipHandler();
    await out.loadFromBuffer(buf2);
    expect(out.hasFile('word/charts/chart1.xml')).toBe(true);
    expect(out.hasFile('word/diagrams/data1.xml')).toBe(true);
    expect(out.hasFile('word/embeddings/oleObject1.bin')).toBe(true);
  });
});
