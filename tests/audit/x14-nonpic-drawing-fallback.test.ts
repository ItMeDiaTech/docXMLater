/**
 * w:drawing elements whose a:graphicData is not pic:pic — charts (c:chart),
 * SmartArt (dgm:relIds), inline shapes — have no ImageRun model. The parser
 * must fall back to raw-XML preservation instead of silently discarding the
 * run: dropping the body reference orphans the chart/diagram part and the
 * graphic visually disappears from the document.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

const WP_NS = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
const A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DGM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

const CHART_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL_BASE}">` +
  `<c:chart><c:plotArea><c:layout/>` +
  `<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/>` +
  `<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
  `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="2"/></c:catAx>` +
  `<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="1"/></c:valAx>` +
  `</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;

const DIAGRAM_DATA_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<dgm:dataModel xmlns:dgm="${DGM_NS}">` +
  `<dgm:ptLst><dgm:pt modelId="{00000000-0000-0000-0000-000000000001}" type="doc"/></dgm:ptLst>` +
  `</dgm:dataModel>`;

const DIAGRAM_LAYOUT_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<dgm:layoutDef xmlns:dgm="${DGM_NS}" uniqueId="urn:test/layout"><dgm:layoutNode/></dgm:layoutDef>`;

const DIAGRAM_QUICKSTYLE_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<dgm:styleDef xmlns:dgm="${DGM_NS}" uniqueId="urn:test/quickstyle"><dgm:styleLbl name="node0"/></dgm:styleDef>`;

const DIAGRAM_COLORS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<dgm:colorsDef xmlns:dgm="${DGM_NS}" uniqueId="urn:test/colors"><dgm:styleLbl name="node0"/></dgm:colorsDef>`;

const CHART_PARAGRAPH =
  `<w:p><w:r><w:drawing>` +
  `<wp:inline distT="0" distB="0" distL="0" distR="0" ${WP_NS}>` +
  `<wp:extent cx="5274310" cy="3076575"/>` +
  `<wp:docPr id="11" name="Chart 11"/>` +
  `<a:graphic ${A_NS}>` +
  `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
  `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ${R_NS} r:id="rId600"/>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

const SMARTART_PARAGRAPH =
  `<w:p><w:r><w:drawing>` +
  `<wp:inline distT="0" distB="0" distL="0" distR="0" ${WP_NS}>` +
  `<wp:extent cx="5486400" cy="3200400"/>` +
  `<wp:docPr id="12" name="Diagram 12"/>` +
  `<a:graphic ${A_NS}>` +
  `<a:graphicData uri="${DGM_NS}">` +
  `<dgm:relIds xmlns:dgm="${DGM_NS}" ${R_NS} r:dm="rId601" r:lo="rId602" r:qs="rId603" r:cs="rId604"/>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

const INSERTED_CHART_PARAGRAPH =
  `<w:p>` +
  `<w:ins w:id="900" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">` +
  `<w:r><w:drawing>` +
  `<wp:inline distT="0" distB="0" distL="0" distR="0" ${WP_NS}>` +
  `<wp:extent cx="5274310" cy="3076575"/>` +
  `<wp:docPr id="13" name="Chart 13"/>` +
  `<a:graphic ${A_NS}>` +
  `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
  `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ${R_NS} r:id="rId600"/>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>` +
  `</w:ins>` +
  `</w:p>`;

async function buildDocx(paragraphXml: string, withDiagram: boolean): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('Non-picture drawing round-trip');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);

  // Insert the drawing reference into the document body (before sectPr).
  const docXml = zip.getFileAsString('word/document.xml')!;
  const updatedDoc = docXml.includes('<w:sectPr')
    ? docXml.replace('<w:sectPr', `${paragraphXml}<w:sectPr`)
    : docXml.replace('</w:body>', `${paragraphXml}</w:body>`);
  zip.updateFile('word/document.xml', updatedDoc);

  // Wire the parts, relationships, and Content_Types backing the reference.
  zip.addFile('word/charts/chart1.xml', CHART_XML);
  let ctOverrides = `<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`;
  let rels = `<Relationship Id="rId600" Type="${REL_BASE}/chart" Target="charts/chart1.xml"/>`;

  if (withDiagram) {
    zip.addFile('word/diagrams/data1.xml', DIAGRAM_DATA_XML);
    zip.addFile('word/diagrams/layout1.xml', DIAGRAM_LAYOUT_XML);
    zip.addFile('word/diagrams/quickStyle1.xml', DIAGRAM_QUICKSTYLE_XML);
    zip.addFile('word/diagrams/colors1.xml', DIAGRAM_COLORS_XML);
    ctOverrides +=
      `<Override PartName="/word/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>` +
      `<Override PartName="/word/diagrams/layout1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/>` +
      `<Override PartName="/word/diagrams/quickStyle1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"/>` +
      `<Override PartName="/word/diagrams/colors1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/>`;
    rels +=
      `<Relationship Id="rId601" Type="${REL_BASE}/diagramData" Target="diagrams/data1.xml"/>` +
      `<Relationship Id="rId602" Type="${REL_BASE}/diagramLayout" Target="diagrams/layout1.xml"/>` +
      `<Relationship Id="rId603" Type="${REL_BASE}/diagramQuickStyle" Target="diagrams/quickStyle1.xml"/>` +
      `<Relationship Id="rId604" Type="${REL_BASE}/diagramColors" Target="diagrams/colors1.xml"/>`;
  }

  const ct = zip.getFileAsString('[Content_Types].xml')!;
  zip.updateFile('[Content_Types].xml', ct.replace('</Types>', `${ctOverrides}</Types>`));

  const docRels = zip.getFileAsString('word/_rels/document.xml.rels')!;
  zip.updateFile(
    'word/_rels/document.xml.rels',
    docRels.replace('</Relationships>', `${rels}</Relationships>`)
  );

  return zip.toBuffer();
}

async function roundTrip(buf: Buffer): Promise<string> {
  const doc = await Document.loadFromBuffer(buf);
  try {
    const out = new ZipHandler();
    await out.loadFromBuffer(await doc.toBuffer());
    return out.getFileAsString('word/document.xml')!;
  } finally {
    doc.dispose();
  }
}

describe('Non-picture w:drawing (chart/SmartArt) raw-XML fallback', () => {
  it('preserves a chart drawing reference in the body on unmodified round-trip', async () => {
    const docXml = await roundTrip(await buildDocx(CHART_PARAGRAPH, false));
    expect(docXml).toContain('<w:drawing>');
    expect(docXml).toContain(
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    );
    expect(docXml).toMatch(/<c:chart[^>]*r:id="rId600"/);
  });

  it('preserves a SmartArt drawing reference in the body on unmodified round-trip', async () => {
    const docXml = await roundTrip(await buildDocx(SMARTART_PARAGRAPH, true));
    expect(docXml).toContain('<w:drawing>');
    expect(docXml).toContain(`<a:graphicData uri="${DGM_NS}">`);
    expect(docXml).toMatch(/<dgm:relIds[^>]*r:dm="rId601"/);
  });

  it('preserves a chart drawing inside a tracked insertion when revisions are preserved', async () => {
    const buf = await buildDocx(INSERTED_CHART_PARAGRAPH, false);
    const doc = await Document.loadFromBuffer(buf, { revisionHandling: 'preserve' });
    try {
      const out = new ZipHandler();
      await out.loadFromBuffer(await doc.toBuffer());
      const docXml = out.getFileAsString('word/document.xml')!;
      expect(docXml).toContain('<w:drawing>');
      expect(docXml).toMatch(/<c:chart[^>]*r:id="rId600"/);
    } finally {
      doc.dispose();
    }
  });

  it('keeps the chart drawing when the tracked insertion is accepted on load', async () => {
    const docXml = await roundTrip(await buildDocx(INSERTED_CHART_PARAGRAPH, false));
    expect(docXml).toContain('<w:drawing>');
    expect(docXml).toMatch(/<c:chart[^>]*r:id="rId600"/);
  });
});
