/**
 * Parsed hyperlinks must not gain hard-coded direct formatting on round-trip.
 *
 * Word-authored hyperlinks typically carry only <w:rStyle w:val="Hyperlink"/>
 * (or no rPr at all). The Hyperlink constructor's default styling
 * (Verdana 12pt blue underline) is intended for programmatically created
 * links only; merging it under parser-supplied formatting injected direct
 * formatting that never existed in the source and overrode the document's
 * Hyperlink character style on every load→save.
 *
 * These tests pin the fixed behavior: a plain load→save round-trip keeps
 * the hyperlink run's rPr exactly as authored, while API-created hyperlinks
 * still receive the documented default styling.
 */

import { Document } from '../../src/core/Document';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { ZipHandler } from '../../src/zip/ZipHandler';

/**
 * Builds a minimal DOCX whose body is a single paragraph containing one
 * external hyperlink (rId5) with the given inner run XML.
 */
async function buildDocxWithHyperlinkRun(runXml: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();

  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`
  );

  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`
  );

  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5" w:history="1">
        ${runXml}
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`
  );

  zipHandler.addFile(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
</Relationships>`
  );

  zipHandler.addFile(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults/>
</w:styles>`
  );

  zipHandler.addFile(
    'docProps/core.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:creator>Test</dc:creator>
</cp:coreProperties>`
  );

  zipHandler.addFile(
    'docProps/app.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Test</Application>
</Properties>`
  );

  return await zipHandler.toBuffer();
}

/** Extracts the first <w:hyperlink>…</w:hyperlink> fragment from saved document.xml. */
async function saveAndGetHyperlinkXml(doc: Document): Promise<string> {
  const buffer = await doc.toBuffer();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  const content = zip.getFile('word/document.xml')?.content;
  const documentXml = content instanceof Buffer ? content.toString('utf8') : String(content);
  const fragment = documentXml.match(/<w:hyperlink[\s\S]*?<\/w:hyperlink>/)?.[0];
  expect(fragment).toBeDefined();
  return fragment!;
}

describe('Parsed hyperlink round-trip formatting (no injected defaults)', () => {
  it('keeps a run with only <w:rStyle w:val="Hyperlink"/> free of direct formatting', async () => {
    const buffer = await buildDocxWithHyperlinkRun(
      `<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Example link</w:t></w:r>`
    );

    const doc = await Document.loadFromBuffer(buffer);
    try {
      const hyperlinkXml = await saveAndGetHyperlinkXml(doc);

      // Source rPr and text survive
      expect(hyperlinkXml).toContain('w:rStyle w:val="Hyperlink"');
      expect(hyperlinkXml).toContain('Example link');

      // No fabricated direct formatting
      expect(hyperlinkXml).not.toContain('Verdana');
      expect(hyperlinkXml).not.toContain('<w:color w:val="0000FF"/>');
      expect(hyperlinkXml).not.toContain('<w:sz w:val="24"/>');
      expect(hyperlinkXml).not.toContain('<w:u ');
    } finally {
      doc.dispose();
    }
  });

  it('keeps a run with no rPr at all free of direct formatting', async () => {
    const buffer = await buildDocxWithHyperlinkRun(`<w:r><w:t>Plain link</w:t></w:r>`);

    const doc = await Document.loadFromBuffer(buffer);
    try {
      const hyperlinkXml = await saveAndGetHyperlinkXml(doc);

      expect(hyperlinkXml).toContain('Plain link');
      expect(hyperlinkXml).not.toContain('Verdana');
      expect(hyperlinkXml).not.toContain('<w:color w:val="0000FF"/>');
      expect(hyperlinkXml).not.toContain('<w:sz w:val="24"/>');
      expect(hyperlinkXml).not.toContain('<w:u ');
    } finally {
      doc.dispose();
    }
  });

  it('preserves explicit source formatting without merging extra defaults', async () => {
    const buffer = await buildDocxWithHyperlinkRun(
      `<w:r><w:rPr><w:b/><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>Styled link</w:t></w:r>`
    );

    const doc = await Document.loadFromBuffer(buffer);
    try {
      const hyperlink = doc.getParagraphs()[0]!.getContent()[0] as Hyperlink;
      const formatting = hyperlink.getRawFormatting();

      // Exactly what the source declared
      expect(formatting.bold).toBe(true);
      expect(formatting.color).toBe('0563C1');
      expect(formatting.underline).toBe('single');

      // No defaults merged underneath
      expect(formatting.font).toBeUndefined();
      expect(formatting.size).toBeUndefined();

      const hyperlinkXml = await saveAndGetHyperlinkXml(doc);
      expect(hyperlinkXml).toContain('<w:color w:val="0563C1"/>');
      expect(hyperlinkXml).not.toContain('Verdana');
      expect(hyperlinkXml).not.toContain('<w:sz w:val="24"/>');
    } finally {
      doc.dispose();
    }
  });

  it('parsed in-memory formatting carries only the source rPr keys', async () => {
    const buffer = await buildDocxWithHyperlinkRun(
      `<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Example link</w:t></w:r>`
    );

    const doc = await Document.loadFromBuffer(buffer);
    try {
      const hyperlink = doc.getParagraphs()[0]!.getContent()[0] as Hyperlink;
      const formatting = hyperlink.getRawFormatting();

      expect(formatting.characterStyle).toBe('Hyperlink');
      expect(formatting.font).toBeUndefined();
      expect(formatting.size).toBeUndefined();
      expect(formatting.color).toBeUndefined();
      expect(formatting.underline).toBeUndefined();
    } finally {
      doc.dispose();
    }
  });

  it('still applies the documented default styling to API-created hyperlinks', () => {
    const link = Hyperlink.createExternal('https://example.com', 'Link');
    const formatting = link.getRawFormatting();

    expect(formatting.font).toBe('Verdana');
    expect(formatting.size).toBe(12);
    expect(formatting.color).toBe('0000FF');
    expect(formatting.underline).toBe('single');
  });
});
