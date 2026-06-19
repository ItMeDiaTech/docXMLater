/**
 * stripTrackedChanges — self-closing revision markers must be removed
 * before the block-removal regexes run.
 *
 * Per ECMA-376 Part 1 §17.13.5.15, a tracked paragraph-mark deletion is a
 * self-closing `<w:del .../>` inside `w:pPr/w:rPr` that precedes the
 * paired `<w:del>...</w:del>` run wrappers of the same paragraph. The
 * deletion block regex's `[^>]*` also matched the trailing '/' of the
 * self-closing form, so the marker acted as an opening tag and everything
 * up to the next `</w:del>` was consumed — eating `</w:rPr></w:pPr>` plus
 * any intervening untracked content. The self-closing cleanup originally
 * ran last (after the damage), corrupting every revision-carrying part
 * loaded with `revisionHandling: 'strip'`.
 *
 * Per §17.13.5.14, a self-closing `<w:del/>` inside `<w:trPr>` marks the
 * ENTIRE row as deleted; stripping just the marker would resurrect the
 * row, so the whole `<w:tr>` must be removed.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { stripTrackedChanges } from '../../src/processors/stripTrackedChanges';

/** Count opening vs closing tags for one element name (self-closing excluded). */
function expectBalanced(xml: string, tag: string): void {
  const withoutSelfClosing = xml.replace(/<[^<>]*\/>/g, '');
  const opens = (withoutSelfClosing.match(new RegExp(`<${tag}[\\s>]`, 'g')) ?? []).length;
  const closes = (withoutSelfClosing.match(new RegExp(`</${tag}>`, 'g')) ?? []).length;
  expect({ tag, opens, closes }).toEqual({ tag, opens, closes: opens });
}

function makeZipWithDocumentXml(bodyXml: string): ZipHandler {
  const zip = new ZipHandler();
  zip.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  zip.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zip.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${bodyXml}</w:body>
</w:document>`
  );
  return zip;
}

const trackedParagraphDeletion =
  '<w:p>' +
  '<w:pPr><w:rPr><w:del w:id="10" w:author="A" w:date="2026-01-15T10:00:00Z"/></w:rPr></w:pPr>' +
  '<w:r><w:t>untracked kept text</w:t></w:r>' +
  '<w:del w:id="11" w:author="A" w:date="2026-01-15T10:00:00Z">' +
  '<w:r><w:delText>gone</w:delText></w:r></w:del>' +
  '</w:p>';

describe('stripTrackedChanges — self-closing <w:del/> markers (paragraph-mark deletions)', () => {
  it('keeps document.xml well-formed and retains untracked content when a paragraph-mark <w:del/> precedes a paired <w:del> block', async () => {
    const zip = makeZipWithDocumentXml(trackedParagraphDeletion);
    await stripTrackedChanges(zip);
    const after = zip.getFileAsString('word/document.xml') ?? '';

    expect(after).toContain('untracked kept text');
    expect(after).not.toContain('gone');
    expect(after).not.toMatch(/<w:del\b/);
    expect(after).toContain('</w:rPr></w:pPr>');
    for (const tag of ['w:pPr', 'w:rPr', 'w:p', 'w:body']) {
      expectBalanced(after, tag);
    }
  });

  it('handles a self-closing <w:moveFrom/> marker preceding a paired <w:moveFrom> block', async () => {
    const body =
      '<w:p>' +
      '<w:pPr><w:rPr><w:moveFrom w:id="20" w:author="A" w:date="2026-01-15T10:00:00Z"/></w:rPr></w:pPr>' +
      '<w:r><w:t>kept</w:t></w:r>' +
      '<w:moveFrom w:id="21" w:author="A" w:date="2026-01-15T10:00:00Z">' +
      '<w:r><w:t>moved away</w:t></w:r></w:moveFrom>' +
      '</w:p>';
    const zip = makeZipWithDocumentXml(body);
    await stripTrackedChanges(zip);
    const after = zip.getFileAsString('word/document.xml') ?? '';

    expect(after).toContain('kept');
    expect(after).not.toContain('moved away');
    expect(after).not.toMatch(/<w:moveFrom\b/);
    for (const tag of ['w:pPr', 'w:rPr', 'w:p']) {
      expectBalanced(after, tag);
    }
  });

  it('survives an end-to-end load with revisionHandling: "strip" without losing untracked content', async () => {
    const zip = makeZipWithDocumentXml(
      trackedParagraphDeletion + '<w:p><w:r><w:t>tail</w:t></w:r></w:p>'
    );
    const buffer = await zip.toBuffer();

    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'strip' });
    try {
      const text = doc
        .getAllParagraphs()
        .map((p) => p.getText())
        .join(' ');
      expect(text).toContain('untracked kept text');
      expect(text).toContain('tail');
      expect(text).not.toContain('gone');
    } finally {
      doc.dispose();
    }
  });
});

describe('stripTrackedChanges — row-level <w:del/> in w:trPr', () => {
  it('removes the entire row marked as a tracked deletion, keeping siblings', async () => {
    const body = `
    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:trPr/>
        <w:tc><w:tcPr/><w:p><w:r><w:t>keep</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:trPr>
          <w:del w:id="1" w:author="A" w:date="2026-01-15T10:00:00Z"/>
        </w:trPr>
        <w:tc><w:tcPr/><w:p><w:r><w:t>row gone</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:trPr/>
        <w:tc><w:tcPr/><w:p><w:r><w:t>tail</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p/>`;
    const zip = makeZipWithDocumentXml(body);
    await stripTrackedChanges(zip);
    const after = zip.getFileAsString('word/document.xml') ?? '';

    expect(after).toContain('keep');
    expect(after).toContain('tail');
    expect(after).not.toContain('row gone');
    expect(after).not.toMatch(/<w:del\b/);
    for (const tag of ['w:tbl', 'w:tr', 'w:tc', 'w:trPr', 'w:p']) {
      expectBalanced(after, tag);
    }
  });
});
