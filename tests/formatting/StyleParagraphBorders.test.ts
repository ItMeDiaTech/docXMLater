/**
 * Style-level pPr — paragraph borders (w:pBdr) round-trip.
 *
 * CT_PPrBase position #9 (w:pBdr §17.3.1.24 — paragraph borders). The
 * style-level parser reads it into `paragraphFormatting.borders`, but
 * `Style.generateParagraphProperties` was silently dropping it on save:
 * any paragraph style with a borders override was lost on programmatic
 * resave that bypassed raw-XML passthrough.
 *
 * Per ECMA-376, paragraph pBdr supports six border sides: top, left,
 * bottom, right, between, bar.
 */

import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyleBorders(pBdrXml: string): Promise<Buffer> {
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
  <w:style w:type="paragraph" w:styleId="BorderTest">
    <w:name w:val="BorderTest"/>
    <w:pPr>${pBdrXml}</w:pPr>
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
  return doc.getStylesManager().getStyle('BorderTest')?.getParagraphFormatting();
}

describe('Style pPr — paragraph borders (§17.3.1.24)', () => {
  it('parses <w:pBdr> with all six border sides', async () => {
    const pBdrXml = `
      <w:pBdr>
        <w:top w:val="single" w:sz="4" w:space="1" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:space="4" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="1" w:color="000000"/>
        <w:right w:val="single" w:sz="4" w:space="4" w:color="000000"/>
        <w:between w:val="single" w:sz="4" w:space="1" w:color="FF0000"/>
        <w:bar w:val="single" w:sz="4" w:space="4" w:color="00FF00"/>
      </w:pBdr>`;
    const buffer = await makeDocxWithStyleBorders(pBdrXml);
    const doc = await Document.loadFromBuffer(buffer);
    const borders = getFmt(doc)?.borders;
    expect(borders).toBeDefined();
    expect(borders?.top?.style).toBe('single');
    expect(borders?.between?.color).toBe('FF0000');
    expect(borders?.bar?.color).toBe('00FF00');
    doc.dispose();
  });

  it('emits <w:pBdr> via Style.toXML() when borders are set', () => {
    const style = new Style({
      styleId: 'BorderTest',
      type: 'paragraph',
      name: 'BorderTest',
      paragraphFormatting: {
        borders: {
          top: { style: 'single', size: 4, space: 1, color: '000000' },
          bottom: { style: 'double', size: 6, space: 2, color: 'FF0000' },
        },
      },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toContain('<w:pBdr>');
    expect(xml).toContain('<w:top');
    expect(xml).toContain('w:val="single"');
    expect(xml).toContain('<w:bottom');
    expect(xml).toContain('w:val="double"');
    expect(xml).toContain('w:color="FF0000"');
    expect(xml).toContain('</w:pBdr>');
  });

  it('emits borders in schema order: top → left → bottom → right → between → bar', () => {
    const style = new Style({
      styleId: 'BorderTest',
      type: 'paragraph',
      name: 'BorderTest',
      paragraphFormatting: {
        borders: {
          bar: { style: 'single', size: 4, color: '000000' },
          between: { style: 'single', size: 4, color: '000000' },
          right: { style: 'single', size: 4, color: '000000' },
          bottom: { style: 'single', size: 4, color: '000000' },
          left: { style: 'single', size: 4, color: '000000' },
          top: { style: 'single', size: 4, color: '000000' },
        },
      },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    const pBdrIdx = xml.indexOf('<w:pBdr>');
    const pBdrBlock = xml.substring(pBdrIdx);
    const topIdx = pBdrBlock.indexOf('<w:top');
    const leftIdx = pBdrBlock.indexOf('<w:left');
    const bottomIdx = pBdrBlock.indexOf('<w:bottom');
    const rightIdx = pBdrBlock.indexOf('<w:right');
    const betweenIdx = pBdrBlock.indexOf('<w:between');
    const barIdx = pBdrBlock.indexOf('<w:bar');
    expect(topIdx).toBeGreaterThan(-1);
    expect(leftIdx).toBeGreaterThan(topIdx);
    expect(bottomIdx).toBeGreaterThan(leftIdx);
    expect(rightIdx).toBeGreaterThan(bottomIdx);
    expect(betweenIdx).toBeGreaterThan(rightIdx);
    expect(barIdx).toBeGreaterThan(betweenIdx);
  });

  it('omits <w:pBdr> when no borders are set', () => {
    const style = new Style({
      styleId: 'BorderTest',
      type: 'paragraph',
      name: 'BorderTest',
      paragraphFormatting: {},
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).not.toContain('<w:pBdr');
  });

  it('emits pBdr between suppressLineNumbers and shd per CT_PPrBase order', () => {
    const style = new Style({
      styleId: 'BorderTest',
      type: 'paragraph',
      name: 'BorderTest',
      paragraphFormatting: {
        suppressLineNumbers: true,
        borders: { top: { style: 'single', size: 4, color: '000000' } },
        shading: { pattern: 'clear', fill: 'FFFF00' },
      },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    const slnIdx = xml.indexOf('<w:suppressLineNumbers');
    const pBdrIdx = xml.indexOf('<w:pBdr');
    const shdIdx = xml.indexOf('<w:shd');
    expect(slnIdx).toBeGreaterThan(-1);
    expect(pBdrIdx).toBeGreaterThan(slnIdx);
    expect(shdIdx).toBeGreaterThan(pBdrIdx);
  });
});
