/**
 * revisionHandling: 'reject' — revert a tracked-changes document to its
 * original pre-edit state.
 *
 * Reject is the exact inverse of accept:
 *   - insertions (w:ins) are removed (inserted content discarded)
 *   - deletions (w:del) are unwrapped and w:delText restored to w:t
 *   - moveFrom is restored at the source, moveTo discarded at the destination
 *   - property changes (w:rPrChange / w:pPrChange / ...) restore the previous
 *     formatting captured inside the change marker
 *   - inserted rows/tables are removed; deleted rows/tables are kept
 *
 * Each fixture is the EDITED document.xml (with tracked changes). Loading with
 * 'reject' must reproduce the original; loading with 'accept' must produce the
 * fully-edited result. The two together prove they are inverses.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Table } from '../../src/elements/Table';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { rejectAllRevisions } from '../../src/processors/acceptRevisions';

const DOC_OPEN =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<w:body>';
const DOC_CLOSE = '<w:sectPr/></w:body></w:document>';

async function makeDocx(bodyXml: string): Promise<Buffer> {
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
  zipHandler.addFile('word/document.xml', DOC_OPEN + bodyXml + DOC_CLOSE);
  return await zipHandler.toBuffer();
}

async function savedXml(doc: Document): Promise<string> {
  const buffer = await doc.toBuffer();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  const xml = zip.getFileAsString('word/document.xml')!;
  return xml;
}

/** Concatenated text of every body-level paragraph. */
function paragraphTexts(doc: Document): string[] {
  return doc
    .getBodyElements()
    .filter((el): el is Paragraph => el instanceof Paragraph)
    .map((p) => p.getText());
}

/**
 * Run the raw-XML reject processor over a hand-authored `<w:body>` inner string
 * (the caller supplies its own trailing `<w:sectPr>` if needed) and return the
 * transformed `word/document.xml`. Used for schema-ordering assertions that
 * would otherwise be obscured by full-document parse/relationship resolution.
 */
async function rejectRawDocumentXml(bodyInnerXml: string): Promise<string> {
  const zip = new ZipHandler();
  zip.addFile(
    'word/document.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<w:body>' +
      bodyInnerXml +
      '</w:body></w:document>'
  );
  await rejectAllRevisions(zip);
  return zip.getFileAsString('word/document.xml')!;
}

describe("revisionHandling: 'reject' — content revisions", () => {
  it('removes inserted content (w:ins) — inverse of accept keeping it', async () => {
    const body =
      '<w:p>' +
      '<w:r><w:t xml:space="preserve">Hello </w:t></w:r>' +
      '<w:ins w:id="1" w:author="A" w:date="2026-01-15T10:00:00Z">' +
      '<w:r><w:t xml:space="preserve">beautiful </w:t></w:r>' +
      '</w:ins>' +
      '<w:r><w:t>world</w:t></w:r>' +
      '</w:p>';
    const buffer = await makeDocx(body);

    const rejected = await Document.loadFromBuffer(buffer, { revisionHandling: 'reject' });
    expect(paragraphTexts(rejected)).toEqual(['Hello world']);
    const xml = await savedXml(rejected);
    expect(xml).not.toContain('<w:ins');
    expect(xml).not.toContain('beautiful');
    rejected.dispose();

    const accepted = await Document.loadFromBuffer(buffer, { revisionHandling: 'accept' });
    expect(paragraphTexts(accepted)).toEqual(['Hello beautiful world']);
    accepted.dispose();
  });

  it('restores deleted content (w:del) and converts w:delText back to w:t', async () => {
    const body =
      '<w:p>' +
      '<w:r><w:t xml:space="preserve">Hello </w:t></w:r>' +
      '<w:del w:id="2" w:author="A" w:date="2026-01-15T10:00:00Z">' +
      '<w:r><w:delText xml:space="preserve">cruel </w:delText></w:r>' +
      '</w:del>' +
      '<w:r><w:t>world</w:t></w:r>' +
      '</w:p>';
    const buffer = await makeDocx(body);

    const rejected = await Document.loadFromBuffer(buffer, { revisionHandling: 'reject' });
    expect(paragraphTexts(rejected)).toEqual(['Hello cruel world']);
    const xml = await savedXml(rejected);
    expect(xml).not.toContain('<w:del');
    expect(xml).not.toContain('delText');
    expect(xml).toContain('cruel ');
    rejected.dispose();

    const accepted = await Document.loadFromBuffer(buffer, { revisionHandling: 'accept' });
    expect(paragraphTexts(accepted)).toEqual(['Hello world']);
    accepted.dispose();
  });

  it('restores a move: keeps the source (moveFrom), drops the destination (moveTo)', async () => {
    const body =
      '<w:p>' +
      '<w:r><w:t xml:space="preserve">Before </w:t></w:r>' +
      '<w:moveFromRangeStart w:id="20" w:name="mv"/>' +
      '<w:moveFrom w:id="21" w:author="A" w:date="2026-01-15T10:00:00Z">' +
      '<w:r><w:t>MOVED</w:t></w:r>' +
      '</w:moveFrom>' +
      '<w:moveFromRangeEnd w:id="20"/>' +
      '<w:r><w:t xml:space="preserve"> after</w:t></w:r>' +
      '</w:p>' +
      '<w:p>' +
      '<w:r><w:t xml:space="preserve">Dest: </w:t></w:r>' +
      '<w:moveToRangeStart w:id="22" w:name="mv"/>' +
      '<w:moveTo w:id="23" w:author="A" w:date="2026-01-15T10:00:00Z">' +
      '<w:r><w:t>MOVED</w:t></w:r>' +
      '</w:moveTo>' +
      '<w:moveToRangeEnd w:id="22"/>' +
      '</w:p>';
    const buffer = await makeDocx(body);

    const rejected = await Document.loadFromBuffer(buffer, { revisionHandling: 'reject' });
    expect(paragraphTexts(rejected)).toEqual(['Before MOVED after', 'Dest: ']);
    const xml = await savedXml(rejected);
    expect(xml).not.toContain('<w:moveFrom');
    expect(xml).not.toContain('<w:moveTo');
    rejected.dispose();

    const accepted = await Document.loadFromBuffer(buffer, { revisionHandling: 'accept' });
    expect(paragraphTexts(accepted)).toEqual(['Before  after', 'Dest: MOVED']);
    accepted.dispose();
  });
});

describe("revisionHandling: 'reject' — property changes", () => {
  it('restores previous run formatting from w:rPrChange', async () => {
    // Current run is bold; the change records that it was previously italic.
    const body =
      '<w:p>' +
      '<w:r>' +
      '<w:rPr>' +
      '<w:b/>' +
      '<w:rPrChange w:id="3" w:author="A" w:date="2026-01-15T10:00:00Z">' +
      '<w:rPr><w:i/></w:rPr>' +
      '</w:rPrChange>' +
      '</w:rPr>' +
      '<w:t>styled</w:t>' +
      '</w:r>' +
      '</w:p>';
    const buffer = await makeDocx(body);

    const rejected = await Document.loadFromBuffer(buffer, { revisionHandling: 'reject' });
    const rRun = (rejected.getBodyElements()[0] as Paragraph).getRuns()[0]!;
    expect(rRun.getItalic()).toBe(true);
    expect(rRun.getBold()).toBe(false);
    const xml = await savedXml(rejected);
    expect(xml).not.toContain('rPrChange');
    rejected.dispose();

    const accepted = await Document.loadFromBuffer(buffer, { revisionHandling: 'accept' });
    const aRun = (accepted.getBodyElements()[0] as Paragraph).getRuns()[0]!;
    expect(aRun.getBold()).toBe(true);
    expect(aRun.getItalic()).toBe(false);
    accepted.dispose();
  });

  it('restores previous paragraph formatting from w:pPrChange while keeping the run', async () => {
    // Current alignment is center; the change records it was previously left.
    const body =
      '<w:p>' +
      '<w:pPr>' +
      '<w:jc w:val="center"/>' +
      '<w:pPrChange w:id="4" w:author="A" w:date="2026-01-15T10:00:00Z">' +
      '<w:pPr><w:jc w:val="left"/></w:pPr>' +
      '</w:pPrChange>' +
      '</w:pPr>' +
      '<w:r><w:t>aligned</w:t></w:r>' +
      '</w:p>';
    const buffer = await makeDocx(body);

    const rejected = await Document.loadFromBuffer(buffer, { revisionHandling: 'reject' });
    const rPara = rejected.getBodyElements()[0] as Paragraph;
    expect(rPara.getAlignment()).toBe('left');
    expect(rPara.getText()).toBe('aligned');
    const xml = await savedXml(rejected);
    expect(xml).not.toContain('pPrChange');
    rejected.dispose();

    const accepted = await Document.loadFromBuffer(buffer, { revisionHandling: 'accept' });
    const aPara = accepted.getBodyElements()[0] as Paragraph;
    expect(aPara.getAlignment()).toBe('center');
    accepted.dispose();
  });
});

describe("revisionHandling: 'reject' — table rows", () => {
  it('removes inserted rows and keeps deleted rows (inverse of accept)', async () => {
    // Original table had two rows: "keep" and "wasDeleted". The edit deleted
    // the second row and inserted a new "wasInserted" row.
    const body =
      '<w:tbl>' +
      '<w:tblPr/>' +
      '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>' +
      '<w:tr><w:trPr/><w:tc><w:tcPr/><w:p><w:r><w:t>keep</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:trPr><w:del w:id="5" w:author="A" w:date="2026-01-15T10:00:00Z"/></w:trPr>' +
      '<w:tc><w:tcPr/><w:p><w:r><w:t>wasDeleted</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:trPr><w:ins w:id="6" w:author="A" w:date="2026-01-15T10:00:00Z"/></w:trPr>' +
      '<w:tc><w:tcPr/><w:p><w:r><w:t>wasInserted</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>' +
      '<w:p/>';
    const buffer = await makeDocx(body);

    const rejected = await Document.loadFromBuffer(buffer, { revisionHandling: 'reject' });
    const rTable = rejected.getBodyElements().find((el) => el instanceof Table) as Table;
    const rTexts = rTable.getRows().map((r) => r.getCells()[0]!.getParagraphs()[0]!.getText());
    expect(rTexts).toEqual(['keep', 'wasDeleted']);
    const xml = await savedXml(rejected);
    expect(xml).not.toContain('<w:ins');
    expect(xml).not.toContain('<w:del');
    rejected.dispose();

    const accepted = await Document.loadFromBuffer(buffer, { revisionHandling: 'accept' });
    const aTable = accepted.getBodyElements().find((el) => el instanceof Table) as Table;
    const aTexts = aTable.getRows().map((r) => r.getCells()[0]!.getParagraphs()[0]!.getText());
    expect(aTexts).toEqual(['keep', 'wasInserted']);
    accepted.dispose();
  });
});

describe("revisionHandling: 'reject' — invariants", () => {
  it('leaves no revision markup after rejecting a mixed document', async () => {
    const body =
      '<w:p>' +
      '<w:ins w:id="1" w:author="A" w:date="2026-01-15T10:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins>' +
      '<w:del w:id="2" w:author="A" w:date="2026-01-15T10:00:00Z"><w:r><w:delText>removed</w:delText></w:r></w:del>' +
      '<w:r>' +
      '<w:rPr><w:b/><w:rPrChange w:id="3" w:author="A" w:date="2026-01-15T10:00:00Z"><w:rPr/></w:rPrChange></w:rPr>' +
      '<w:t>kept</w:t>' +
      '</w:r>' +
      '</w:p>';
    const buffer = await makeDocx(body);

    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'reject' });
    const xml = await savedXml(doc);
    for (const marker of [
      '<w:ins',
      '<w:del',
      '<w:moveFrom',
      '<w:moveTo',
      'delText',
      'rPrChange',
      'pPrChange',
    ]) {
      expect(xml).not.toContain(marker);
    }
    // Inserted content gone, deleted content restored.
    expect(paragraphTexts(doc)).toEqual(['removedkept']);
    // The run that had a w:rPrChange with an empty previous-rPr loses its bold.
    const run = (doc.getBodyElements()[0] as Paragraph)
      .getRuns()
      .find((r) => r.getText() === 'kept')!;
    expect(run.getBold()).toBe(false);
    doc.dispose();
  });

  it('is a no-op for a document with no tracked changes', async () => {
    const body = '<w:p><w:r><w:t>plain text</w:t></w:r></w:p>';
    const buffer = await makeDocx(body);
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'reject' });
    expect(paragraphTexts(doc)).toEqual(['plain text']);
    doc.dispose();
  });
});

describe("revisionHandling: 'reject' — property-change schema fidelity", () => {
  it('rejecting w:sectPrChange preserves header/footer references (prepended) and reverts the base props', async () => {
    const xml = await rejectRawDocumentXml(
      '<w:p><w:r><w:t>x</w:t></w:r></w:p>' +
        '<w:sectPr>' +
        '<w:headerReference w:type="default" r:id="rId10"/>' +
        '<w:footerReference w:type="default" r:id="rId11"/>' +
        '<w:pgSz w:w="12240" w:h="15840"/>' +
        '<w:sectPrChange w:id="7" w:author="A" w:date="2026-01-15T10:00:00Z">' +
        '<w:sectPr><w:pgSz w:w="15840" w:h="12240"/></w:sectPr>' +
        '</w:sectPrChange>' +
        '</w:sectPr>'
    );
    // References survive (the snapshot is CT_SectPrBase and cannot carry them).
    expect(xml).toContain('w:headerReference');
    expect(xml).toContain('w:footerReference');
    // Base section props reverted to the previous (landscape) page size.
    expect(xml).toContain('w:w="15840"');
    expect(xml).not.toContain('w:w="12240"');
    expect(xml).not.toContain('sectPrChange');
    // CT_SectPr order: header/footer references precede the base contents.
    expect(xml.indexOf('w:headerReference')).toBeLessThan(xml.indexOf('w:pgSz'));
    expect(xml.indexOf('w:footerReference')).toBeLessThan(xml.indexOf('w:pgSz'));
  });

  it('rejecting w:tblGridChange reverts the grid even when it is the only marker in the part', async () => {
    const xml = await rejectRawDocumentXml(
      '<w:tbl><w:tblPr/>' +
        '<w:tblGrid>' +
        '<w:gridCol w:w="5000"/>' +
        '<w:gridCol w:w="5000"/>' +
        '<w:tblGridChange w:id="8">' +
        '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="5320"/></w:tblGrid>' +
        '</w:tblGridChange>' +
        '</w:tblGrid>' +
        '<w:tr><w:trPr/><w:tc><w:tcPr/><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr>' +
        '</w:tbl>'
    );
    expect(xml).toContain('w:w="4680"');
    expect(xml).toContain('w:w="5320"');
    expect(xml).not.toContain('w:w="5000"');
    expect(xml).not.toContain('tblGridChange');
  });

  it('rejecting w:tblPrChange restores the previous table properties', async () => {
    const xml = await rejectRawDocumentXml(
      '<w:tbl>' +
        '<w:tblPr>' +
        '<w:tblW w:w="5000" w:type="dxa"/>' +
        '<w:tblPrChange w:id="9" w:author="A" w:date="2026-01-15T10:00:00Z">' +
        '<w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>' +
        '</w:tblPrChange>' +
        '</w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' +
        '<w:tr><w:trPr/><w:tc><w:tcPr/><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr>' +
        '</w:tbl>'
    );
    expect(xml).toContain('w:w="9000"');
    expect(xml).not.toContain('w:w="5000"');
    expect(xml).not.toContain('tblPrChange');
  });

  it('rejecting a nested pPr>rPr>rPrChange restores both the paragraph and the mark-run formatting', async () => {
    const xml = await rejectRawDocumentXml(
      '<w:p><w:pPr>' +
        '<w:jc w:val="center"/>' +
        '<w:rPr><w:b/>' +
        '<w:rPrChange w:id="12" w:author="A" w:date="2026-01-15T10:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange>' +
        '</w:rPr>' +
        '<w:pPrChange w:id="13" w:author="A" w:date="2026-01-15T10:00:00Z"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>' +
        '</w:pPr>' +
        '<w:r><w:t>x</w:t></w:r></w:p>'
    );
    // Paragraph property reverted, mark-run rPr restored (italic, not bold).
    expect(xml).toContain('w:val="left"');
    expect(xml).not.toContain('w:val="center"');
    expect(xml).toContain('<w:i/>');
    expect(xml).not.toContain('<w:b/>');
    expect(xml).not.toContain('pPrChange');
    expect(xml).not.toContain('rPrChange');
    // The preserved mark rPr stays after the base paragraph property (CT_PPr order).
    expect(xml.indexOf('w:jc')).toBeLessThan(xml.indexOf('<w:rPr>'));
  });

  it('drops a malformed *PrChange with no embedded snapshot, keeping current formatting', async () => {
    const xml = await rejectRawDocumentXml(
      '<w:p><w:r>' +
        '<w:rPr><w:b/><w:rPrChange w:id="14" w:author="A" w:date="2026-01-15T10:00:00Z"/></w:rPr>' +
        '<w:t>kept</w:t>' +
        '</w:r></w:p>'
    );
    // No snapshot to restore => keep the live formatting, just drop the marker.
    expect(xml).toContain('<w:b/>');
    expect(xml).not.toContain('rPrChange');
  });

  it('restores deleted text in original order when a run holds w:delText before w:t', async () => {
    const xml = await rejectRawDocumentXml(
      '<w:p>' +
        '<w:del w:id="2" w:author="A" w:date="2026-01-15T10:00:00Z">' +
        '<w:r><w:delText>B</w:delText><w:t>A</w:t></w:r>' +
        '</w:del>' +
        '</w:p>'
    );
    expect(xml).not.toContain('delText');
    // Document order preserved: B (restored from delText) precedes A.
    expect(xml.indexOf('>B<')).toBeLessThan(xml.indexOf('>A<'));
  });
});

describe("revisionHandling: 'reject' — conflicting options", () => {
  it('throws when acceptRevisions:true is combined with revisionHandling:reject', async () => {
    const buffer = await makeDocx('<w:p><w:r><w:t>x</w:t></w:r></w:p>');
    await expect(
      Document.loadFromBuffer(buffer, { acceptRevisions: true, revisionHandling: 'reject' })
    ).rejects.toThrow(/cannot be combined/i);
  });
});
