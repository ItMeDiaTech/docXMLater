/**
 * `<w:fldSimple>` cached field result round-trip.
 *
 * Per ECMA-376 Part 1 §17.16.16 CT_SimpleField, the child runs of
 * `<w:fldSimple>` hold the field's current (cached) result. Word does
 * not refresh fields on open by default, so that cached snapshot is the
 * text the user actually sees. The parser previously discarded the
 * child runs and the emitter substituted a synthetic placeholder
 * (PAGE → "1", AUTHOR → "Author", REF → "1", unknown → ""), silently
 * corrupting visible content and dropping per-run formatting.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Field } from '../../src/elements/Field';
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
  try {
    const out = await doc.toBuffer();
    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const content = zip.getFile('word/document.xml')?.content;
    return content instanceof Buffer ? content.toString('utf8') : String(content);
  } finally {
    doc.dispose();
  }
}

function buildDoc(fldSimpleXml: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>${fldSimpleXml}</w:p>
  </w:body>
</w:document>`;
}

function extractFldSimple(xml: string): string {
  return xml.match(/<w:fldSimple[\s\S]*?<\/w:fldSimple>/)?.[0] ?? '';
}

describe('<w:fldSimple> cached result round-trip (ECMA-376 §17.16.16)', () => {
  it('preserves a REF field cached result instead of replacing it with "1"', async () => {
    const out = await loadAndResaveDocXml(
      buildDoc(
        '<w:fldSimple w:instr=" REF _Ref12345 \\h "><w:r><w:t>Cached Heading Result</w:t></w:r></w:fldSimple>'
      )
    );
    const fldSimple = extractFldSimple(out);
    expect(fldSimple).toContain('Cached Heading Result');
    expect(fldSimple).not.toMatch(/<w:t[^>]*>1<\/w:t>/);
  });

  it('preserves per-run formatting (w:rPr) on the cached result', async () => {
    const out = await loadAndResaveDocXml(
      buildDoc(
        '<w:fldSimple w:instr=" REF _Ref12345 \\h "><w:r><w:rPr><w:b/></w:rPr><w:t>Bold Result</w:t></w:r></w:fldSimple>'
      )
    );
    const fldSimple = extractFldSimple(out);
    expect(fldSimple).toContain('Bold Result');
    // CT_OnOff true may round-trip as <w:b/> or the explicit <w:b w:val="1"/>
    expect(fldSimple).toMatch(/<w:rPr>[\s\S]*?<w:b\b[^>]*\/>[\s\S]*?<\/w:rPr>/);
  });

  it('preserves an AUTHOR field cached result instead of replacing it with "Author"', async () => {
    const out = await loadAndResaveDocXml(
      buildDoc(
        '<w:fldSimple w:instr=" AUTHOR \\* MERGEFORMAT "><w:r><w:t>Jane Q. Author</w:t></w:r></w:fldSimple>'
      )
    );
    const fldSimple = extractFldSimple(out);
    expect(fldSimple).toContain('Jane Q. Author');
    expect(fldSimple).not.toMatch(/<w:t[^>]*>Author<\/w:t>/);
  });

  it('preserves a PAGE field cached result that differs from the placeholder', async () => {
    const out = await loadAndResaveDocXml(
      buildDoc('<w:fldSimple w:instr=" PAGE "><w:r><w:t>42</w:t></w:r></w:fldSimple>')
    );
    expect(extractFldSimple(out)).toMatch(/<w:t[^>]*>42<\/w:t>/);
  });

  it('preserves multiple cached result runs with distinct formatting', async () => {
    const out = await loadAndResaveDocXml(
      buildDoc(
        '<w:fldSimple w:instr=" REF _Ref99 \\h ">' +
          '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Chapter </w:t></w:r>' +
          '<w:r><w:rPr><w:i/></w:rPr><w:t>Seven</w:t></w:r>' +
          '</w:fldSimple>'
      )
    );
    const fldSimple = extractFldSimple(out);
    // Both runs survive in order with their own rPr
    const runs = fldSimple.match(/<w:r>[\s\S]*?<\/w:r>/g) ?? [];
    expect(runs).toHaveLength(2);
    expect(runs[0]).toMatch(/<w:b\b[^>]*\/>/);
    expect(runs[0]).toMatch(/<w:t[^>]*>Chapter <\/w:t>/);
    expect(runs[1]).toMatch(/<w:i\b[^>]*\/>/);
    expect(runs[1]).toMatch(/<w:t[^>]*>Seven<\/w:t>/);
  });

  it('still emits a placeholder for programmatically created fields', async () => {
    const doc = Document.create();
    try {
      const para = new Paragraph();
      para.addField(Field.createPageNumber());
      doc.addParagraph(para);
      const out = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const content = zip.getFile('word/document.xml')?.content;
      const xml = content instanceof Buffer ? content.toString('utf8') : String(content);
      expect(extractFldSimple(xml)).toMatch(/<w:t[^>]*>1<\/w:t>/);
    } finally {
      doc.dispose();
    }
  });
});
