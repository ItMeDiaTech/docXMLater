/**
 * Style-level pPr — shading (w:shd) round-trip.
 *
 * CT_PPrBase position #10 (w:shd §17.3.1.32 — paragraph shading). The
 * style-level parser reads it into `paragraphFormatting.shading`, but
 * `Style.generateParagraphProperties` was silently dropping it on save:
 * any style with a paragraph-level shading override would be serialized
 * without the shading element, causing the override to be lost on any
 * programmatic modification that bypassed raw-XML passthrough.
 */

import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyleShading(shdInner: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();
  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );
  zipHandler.addFile(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="ShdTest">
    <w:name w:val="ShdTest"/>
    <w:pPr>${shdInner}</w:pPr>
  </w:style>
</w:styles>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>test</w:t></w:r></w:p></w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

function getFmt(doc: Document) {
  return doc.getStylesManager().getStyle('ShdTest')?.getParagraphFormatting();
}

describe('Style pPr — shading (§17.3.1.32)', () => {
  it('parses <w:shd> with fill/pattern/color attrs', async () => {
    const buffer = await makeDocxWithStyleShading(
      '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const shading = getFmt(doc)?.shading;
    expect(shading).toBeDefined();
    expect(shading?.fill).toBe('FFFF00');
    expect(shading?.pattern).toBe('clear');
    expect(shading?.color).toBe('auto');
    doc.dispose();
  });

  it('emits <w:shd> via Style.toXML() when shading is set', () => {
    const style = new Style({
      styleId: 'ShdTest',
      type: 'paragraph',
      name: 'ShdTest',
      paragraphFormatting: {
        shading: { pattern: 'clear', color: 'auto', fill: 'FFFF00' },
      },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toContain('<w:shd');
    expect(xml).toContain('w:fill="FFFF00"');
    expect(xml).toContain('w:val="clear"');
  });

  it('omits <w:shd/> when shading is undefined', () => {
    const style = new Style({
      styleId: 'ShdTest',
      type: 'paragraph',
      name: 'ShdTest',
      paragraphFormatting: {},
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).not.toContain('<w:shd');
  });

  it('emits shd between suppressLineNumbers and suppressAutoHyphens per CT_PPrBase', () => {
    const style = new Style({
      styleId: 'ShdTest',
      type: 'paragraph',
      name: 'ShdTest',
      paragraphFormatting: {
        suppressLineNumbers: true,
        shading: { pattern: 'clear', fill: 'FF0000' },
        suppressAutoHyphens: true,
      },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    const slnIdx = xml.indexOf('<w:suppressLineNumbers');
    const shdIdx = xml.indexOf('<w:shd');
    const sahIdx = xml.indexOf('<w:suppressAutoHyphens');
    expect(slnIdx).toBeGreaterThan(-1);
    expect(shdIdx).toBeGreaterThan(slnIdx);
    expect(sahIdx).toBeGreaterThan(shdIdx);
  });
});
