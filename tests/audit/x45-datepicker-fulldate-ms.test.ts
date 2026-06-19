/**
 * w:fullDate on the date-picker w:date element is ST_DateTime, and the
 * framework's Word-compat standard (src/utils/dateFormatting.ts) is ISO
 * 8601 WITHOUT milliseconds — Word rejects them in w:date attributes.
 * Every other date emission site routes through formatDateForXml();
 * the SDT date picker used Date.toISOString(), which always emits
 * '.sssZ', so a parsed '2024-01-01T00:00:00Z' degraded to
 * '2024-01-01T00:00:00.000Z' whenever sdtPr was rebuilt.
 */
import { Document } from '../../src/core/Document';
import { StructuredDocumentTag } from '../../src/elements/StructuredDocumentTag';
import { Paragraph } from '../../src/elements/Paragraph';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLElement } from '../../src/xml/XMLBuilder';

const NO_MS_ISO = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/;

function findChild(parent: XMLElement, name: string): XMLElement | undefined {
  return parent.children?.find(
    (child): child is XMLElement => typeof child !== 'string' && child.name === name
  );
}

const DATE_SDT =
  `<w:sdt>` +
  `<w:sdtPr>` +
  `<w:id w:val="55"/>` +
  `<w:tag w:val="orig-tag"/>` +
  `<w:date w:fullDate="2024-06-15T00:00:00Z">` +
  `<w:dateFormat w:val="M/d/yyyy"/>` +
  `<w:lid w:val="en-US"/>` +
  `</w:date>` +
  `</w:sdtPr>` +
  `<w:sdtContent><w:p><w:r><w:t>6/15/2024</w:t></w:r></w:p></w:sdtContent>` +
  `</w:sdt>`;

async function buildDocxWithDatePicker(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('fullDate round-trip');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);
  const docXml = zip.getFileAsString('word/document.xml')!;
  const updated = docXml.includes('<w:sectPr')
    ? docXml.replace('<w:sectPr', `${DATE_SDT}<w:sectPr`)
    : docXml.replace('</w:body>', `${DATE_SDT}</w:body>`);
  zip.updateFile('word/document.xml', updated);
  return zip.toBuffer();
}

describe('date picker w:fullDate format (no milliseconds)', () => {
  it('emits w:fullDate without milliseconds for a programmatic date picker', () => {
    const sdt = StructuredDocumentTag.createDatePicker('M/d/yyyy', [
      new Paragraph().addText('1/1/2024'),
    ]);
    sdt.setDatePickerProperties({
      dateFormat: 'M/d/yyyy',
      fullDate: new Date('2024-01-01T00:00:00Z'),
    });

    const sdtPr = findChild(sdt.toXML(), 'w:sdtPr');
    expect(sdtPr).toBeDefined();
    const dateElement = findChild(sdtPr!, 'w:date');
    expect(dateElement).toBeDefined();
    const fullDate = dateElement!.attributes!['w:fullDate'];
    expect(fullDate).toBe('2024-01-01T00:00:00Z');
    expect(fullDate).toMatch(NO_MS_ISO);
  });

  it('round-trips a parsed w:fullDate without introducing milliseconds', async () => {
    const doc = await Document.loadFromBuffer(await buildDocxWithDatePicker());
    try {
      const sdt = doc
        .getBodyElements()
        .find((el): el is StructuredDocumentTag => el instanceof StructuredDocumentTag);
      expect(sdt).toBeDefined();
      expect(sdt!.getDatePickerProperties()?.fullDate?.toISOString()).toBe(
        '2024-06-15T00:00:00.000Z'
      );

      // Invalidate the raw sdtPr passthrough so toXML() rebuilds from the model
      sdt!.setTag('mutated-tag');

      const out = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const docXml = zip.getFileAsString('word/document.xml')!;
      expect(docXml).toContain('w:fullDate="2024-06-15T00:00:00Z"');
      expect(docXml).not.toContain('.000Z');
    } finally {
      doc.dispose();
    }
  });
});
