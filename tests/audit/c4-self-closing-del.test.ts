/**
 * stripRevisionsFromXml — self-closing revision markers must not be
 * treated as opening tags.
 *
 * Per ECMA-376 Part 1 §17.13.5.15, a tracked paragraph-mark deletion is
 * serialized as a self-closing `<w:del .../>` inside `w:pPr/w:rPr`,
 * PRECEDING the paired `<w:del>...</w:del>` run wrappers of the same
 * paragraph. The block-removal regex's `[^>]*` also matched the trailing
 * '/' of the self-closing form, so the marker acted as an opening tag and
 * everything up to the next `</w:del>` was swallowed — removing
 * `</w:rPr></w:pPr>` (and any intervening untracked content) and leaving
 * unbalanced XML. This function runs on every `cell.rawNestedContent`
 * nested table in the default `acceptAllRevisions()` path, so the
 * malformed fragment landed verbatim in document.xml.
 */

import { Document } from '../../src/core/Document';
import { stripRevisionsFromXml } from '../../src/processors/InMemoryRevisionAcceptor';
import { ZipHandler } from '../../src/zip/ZipHandler';

/** Count opening vs closing tags for one element name (self-closing excluded). */
function expectBalanced(xml: string, tag: string): void {
  const withoutSelfClosing = xml.replace(/<[^<>]*\/>/g, '');
  const opens = (withoutSelfClosing.match(new RegExp(`<${tag}[\\s>]`, 'g')) ?? []).length;
  const closes = (withoutSelfClosing.match(new RegExp(`</${tag}>`, 'g')) ?? []).length;
  expect({ tag, opens, closes }).toEqual({ tag, opens, closes: opens });
}

describe('stripRevisionsFromXml — self-closing <w:del/> / <w:moveFrom/> markers', () => {
  it('preserves w:pPr/w:rPr structure when a paragraph-mark <w:del/> precedes a paired <w:del> block', () => {
    const xml =
      '<w:tbl><w:tr><w:tc><w:p>' +
      '<w:pPr><w:rPr><w:del w:id="10" w:author="A" w:date="2026-01-15T10:00:00Z"/></w:rPr></w:pPr>' +
      '<w:del w:id="11" w:author="A" w:date="2026-01-15T10:00:00Z">' +
      '<w:r><w:delText>gone</w:delText></w:r></w:del>' +
      '<w:r><w:t>kept</w:t></w:r>' +
      '</w:p></w:tc></w:tr></w:tbl>';

    const result = stripRevisionsFromXml(xml);

    // The untracked run survives; the deleted run does not.
    expect(result).toContain('<w:r><w:t>kept</w:t></w:r>');
    expect(result).not.toContain('gone');
    expect(result).not.toMatch(/<w:del\b/);

    // The paragraph-mark marker must not swallow </w:rPr></w:pPr>.
    expect(result).toContain('</w:rPr></w:pPr>');
    for (const tag of ['w:pPr', 'w:rPr', 'w:p', 'w:tc', 'w:tr', 'w:tbl']) {
      expectBalanced(result, tag);
    }
  });

  it('preserves untracked content located between the marker and the paired block', () => {
    const xml =
      '<w:p>' +
      '<w:pPr><w:rPr><w:del w:id="20" w:author="A" w:date="2026-01-15T10:00:00Z"/></w:rPr></w:pPr>' +
      '<w:r><w:t>untracked kept text</w:t></w:r>' +
      '<w:del w:id="21" w:author="A" w:date="2026-01-15T10:00:00Z">' +
      '<w:r><w:delText>gone</w:delText></w:r></w:del>' +
      '</w:p>';

    const result = stripRevisionsFromXml(xml);

    expect(result).toContain('untracked kept text');
    expect(result).not.toContain('gone');
    expect(result).not.toMatch(/<w:del\b/);
    for (const tag of ['w:pPr', 'w:rPr', 'w:p']) {
      expectBalanced(result, tag);
    }
  });

  it('handles a self-closing <w:moveFrom/> marker preceding a paired <w:moveFrom> block', () => {
    const xml =
      '<w:p>' +
      '<w:pPr><w:rPr><w:moveFrom w:id="30" w:author="A" w:date="2026-01-15T10:00:00Z"/></w:rPr></w:pPr>' +
      '<w:moveFrom w:id="31" w:author="A" w:date="2026-01-15T10:00:00Z">' +
      '<w:r><w:t>moved away</w:t></w:r></w:moveFrom>' +
      '<w:r><w:t>kept</w:t></w:r>' +
      '</w:p>';

    const result = stripRevisionsFromXml(xml);

    expect(result).toContain('<w:r><w:t>kept</w:t></w:r>');
    expect(result).not.toContain('moved away');
    expect(result).not.toMatch(/<w:moveFrom\b/);
    expect(result).toContain('</w:rPr></w:pPr>');
    for (const tag of ['w:pPr', 'w:rPr', 'w:p']) {
      expectBalanced(result, tag);
    }
  });
});

describe('acceptAllRevisions — nested table with tracked paragraph deletion', () => {
  async function makeDocxWithNestedTableDeletion(): Promise<Buffer> {
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
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr/>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
            <w:tr>
              <w:tc>
                <w:tcPr/>
                <w:p><w:pPr><w:rPr><w:del w:id="10" w:author="A" w:date="2026-01-15T10:00:00Z"/></w:rPr></w:pPr><w:del w:id="11" w:author="A" w:date="2026-01-15T10:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>
                <w:p><w:r><w:t>nested kept</w:t></w:r></w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
          <w:p/>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p/>
  </w:body>
</w:document>`
    );
    return await zipHandler.toBuffer();
  }

  it('produces well-formed document.xml after accepting a tracked paragraph deletion inside a nested table', async () => {
    const buffer = await makeDocxWithNestedTableDeletion();
    // 'preserve' keeps revision markup so the in-memory acceptor (the
    // default acceptAllRevisions() path) processes the nested raw XML.
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    try {
      doc.acceptAllRevisions();
      const saved = await doc.toBuffer();

      const outZip = new ZipHandler();
      await outZip.loadFromBuffer(saved);
      const savedXml = outZip.getFileAsString('word/document.xml') ?? '';

      expect(savedXml).toContain('nested kept');
      expect(savedXml).not.toMatch(/<w:del\b/);
      for (const tag of ['w:pPr', 'w:rPr', 'w:p', 'w:tc', 'w:tr', 'w:tbl']) {
        expectBalanced(savedXml, tag);
      }
    } finally {
      doc.dispose();
    }
  });
});
