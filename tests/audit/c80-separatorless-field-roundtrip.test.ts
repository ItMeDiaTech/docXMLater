/**
 * Separator-less complex field round-trip.
 *
 * Per ECMA-376, complex fields without a result section (typical for
 * XE/TC markers and never-updated fields) omit the w:fldChar "separate"
 * element entirely. The parser records this via hasResult=false, and
 * serialization must not synthesize a separator — doing so adds an empty
 * result section the original field never had, breaking round-trip
 * structural fidelity on the first save.
 */

import { Document } from '../../src/core/Document';
import { ComplexField } from '../../src/elements/Field';
import { XMLElement } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

function fldCharTypes(runs: XMLElement[]): unknown[] {
  return runs.flatMap((run) =>
    (run.children || [])
      .filter(
        (child): child is XMLElement => typeof child !== 'string' && child.name === 'w:fldChar'
      )
      .map((child) => child.attributes!['w:fldCharType'])
  );
}

describe('ComplexField separator emission honors hasResult (ECMA-376 round-trip)', () => {
  it('omits the separator for a field without a result section', () => {
    const field = new ComplexField({
      instruction: ' XE "Widget" ',
      hasResult: false,
    });

    const runs = field.toXML();

    expect(fldCharTypes(runs)).toEqual(['begin', 'end']);
    expect(runs).toHaveLength(3);
  });

  it('still emits the separator when the field had one with an empty result', () => {
    const field = new ComplexField({
      instruction: ' PAGE ',
      hasResult: true,
    });

    expect(fldCharTypes(field.toXML())).toEqual(['begin', 'separate', 'end']);
  });

  it('still emits the separator when a result is present', () => {
    const field = new ComplexField({
      instruction: ' PAGE ',
      result: '1',
    });

    expect(fldCharTypes(field.toXML())).toEqual(['begin', 'separate', 'end']);
  });

  it('does not add a separator to a parsed separator-less field on save', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> XE "Widget" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>`;

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
    zipHandler.addFile('word/document.xml', documentXml);
    const buffer = await zipHandler.toBuffer();

    const doc = await Document.loadFromBuffer(buffer);
    try {
      const out = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const content = zip.getFile('word/document.xml')?.content;
      const xml = content instanceof Buffer ? content.toString('utf8') : String(content);

      expect(xml).toContain('w:fldCharType="begin"');
      expect(xml).toContain('w:fldCharType="end"');
      expect(xml).toContain('XE "Widget"');
      expect(xml).not.toContain('w:fldCharType="separate"');
    } finally {
      doc.dispose();
    }
  });
});
