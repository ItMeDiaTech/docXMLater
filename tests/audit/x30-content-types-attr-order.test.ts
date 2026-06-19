/**
 * [Content_Types].xml is regenerated on every save; entries the generator
 * does not emit itself (charts, diagrams, OLE defaults) survive only via the
 * parsed-at-load original entries. Per XML 1.0, attribute order is
 * insignificant, values may be single- or double-quoted, and empty elements
 * may use the expanded form (<Override ...></Override>). Spec-valid packages
 * from non-Word producers must not lose Default/Override entries on
 * round-trip just because they don't match Word's canonical serialization.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

const CHART_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
  `<c:chart><c:plotArea><c:layout/></c:plotArea></c:chart>` +
  `</c:chartSpace>`;

const OLE_BIN = Buffer.from('OLE_BINARY_PAYLOAD', 'utf8');

const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const OLE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.oleObject';

/**
 * Builds a docx whose [Content_Types].xml uses spec-valid but non-Word
 * serialization: a Default with ContentType-before-Extension order in single
 * quotes, and an expanded-form (non-self-closing) Override for the chart part.
 */
async function buildDocxWithNonCanonicalContentTypes(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('Content types attribute order test');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);

  zip.addFile('word/charts/chart1.xml', CHART_XML);
  zip.addFile('word/embeddings/oleObject1.bin', OLE_BIN, { binary: true });

  const ct = zip.getFileAsString('[Content_Types].xml')!;
  const updatedCt = ct.replace(
    '</Types>',
    `<Default ContentType='${OLE_CONTENT_TYPE}' Extension='bin'/>` +
      `<Override PartName="/word/charts/chart1.xml" ContentType="${CHART_CONTENT_TYPE}"></Override>` +
      `</Types>`
  );
  zip.updateFile('[Content_Types].xml', updatedCt);

  return zip.toBuffer();
}

describe('X30: [Content_Types].xml parsing is attribute-order and quote-style independent', () => {
  it('preserves a single-quoted Default with reversed attribute order on round-trip', async () => {
    const buf1 = await buildDocxWithNonCanonicalContentTypes();
    const doc = await Document.loadFromBuffer(buf1);
    try {
      const buf2 = await doc.toBuffer();
      const out = new ZipHandler();
      await out.loadFromBuffer(buf2);
      const ct = out.getFileAsString('[Content_Types].xml')!;

      const binDefault = ct.match(/<Default\b[^>]*Extension="bin"[^>]*>/);
      expect(binDefault).not.toBeNull();
      expect(binDefault![0]).toContain(`ContentType="${OLE_CONTENT_TYPE}"`);
    } finally {
      doc.dispose();
    }
  });

  it('preserves an expanded-form (non-self-closing) Override on round-trip', async () => {
    const buf1 = await buildDocxWithNonCanonicalContentTypes();
    const doc = await Document.loadFromBuffer(buf1);
    try {
      const buf2 = await doc.toBuffer();
      const out = new ZipHandler();
      await out.loadFromBuffer(buf2);
      const ct = out.getFileAsString('[Content_Types].xml')!;

      const chartOverride = ct.match(
        /<Override\b[^>]*PartName="\/word\/charts\/chart1\.xml"[^>]*>/
      );
      expect(chartOverride).not.toBeNull();
      expect(chartOverride![0]).toContain(`ContentType="${CHART_CONTENT_TYPE}"`);
    } finally {
      doc.dispose();
    }
  });

  it('still parses Word-canonical self-closing double-quoted entries', async () => {
    const seed = Document.create();
    seed.createParagraph('Canonical entries');
    const base = await seed.toBuffer();
    seed.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(base);
    zip.addFile('word/charts/chart1.xml', CHART_XML);
    const ct = zip.getFileAsString('[Content_Types].xml')!;
    zip.updateFile(
      '[Content_Types].xml',
      ct.replace(
        '</Types>',
        `<Override PartName="/word/charts/chart1.xml" ContentType="${CHART_CONTENT_TYPE}"/></Types>`
      )
    );
    const buf1 = await zip.toBuffer();

    const doc = await Document.loadFromBuffer(buf1);
    try {
      const buf2 = await doc.toBuffer();
      const out = new ZipHandler();
      await out.loadFromBuffer(buf2);
      const savedCt = out.getFileAsString('[Content_Types].xml')!;
      expect(savedCt).toContain(CHART_CONTENT_TYPE);
    } finally {
      doc.dispose();
    }
  });
});
