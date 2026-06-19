/**
 * Regression tests: parseStyle must capture w:basedOn and w:next from the
 * self-closing form (`<w:basedOn w:val="Normal"/>`). CT_String elements are
 * empty-content per ECMA-376 §17.7.4.3, so Word always serializes them
 * self-closing; parsing them with a closing-tag-only extractor silently
 * strips the style inheritance chain whenever a loaded style is regenerated.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Style } from '../../src/formatting/Style';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading1Char"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
</w:styles>`;

async function buildDocxWithCustomStyles(): Promise<Buffer> {
  const doc = Document.create();
  try {
    doc.addParagraph(new Paragraph().addText('Test'));
    const buffer = await doc.toBuffer();
    const zipHandler = new ZipHandler();
    await zipHandler.loadFromBuffer(buffer);
    zipHandler.updateFile(DOCX_PATHS.STYLES, STYLES_XML);
    return await zipHandler.toBuffer();
  } finally {
    doc.dispose();
  }
}

describe('parseStyle self-closing basedOn/next (X16)', () => {
  it('parses basedOn and next from self-closing elements', async () => {
    const buffer = await buildDocxWithCustomStyles();
    const doc = await Document.loadFromBuffer(buffer);
    try {
      const heading1 = doc.getStyle('Heading1');
      expect(heading1).toBeDefined();
      const props = heading1!.getProperties();
      expect(props.basedOn).toBe('Normal');
      expect(props.next).toBe('Normal');
    } finally {
      doc.dispose();
    }
  });

  it('keeps basedOn/next in saved styles.xml after modifying a loaded style', async () => {
    const buffer = await buildDocxWithCustomStyles();
    const doc = await Document.loadFromBuffer(buffer);
    try {
      // Formatting-only replacement: structural properties must survive the
      // parse -> addStyle -> merge round-trip
      doc.addStyle(
        Style.create({
          styleId: 'Heading1',
          name: 'heading 1',
          type: 'paragraph',
          runFormatting: { bold: true, size: 14 },
        })
      );

      const outputBuffer = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const mergedStylesXml = zip.getFileAsString(DOCX_PATHS.STYLES) || '';

      expect(mergedStylesXml).toContain('<w:basedOn w:val="Normal"/>');
      expect(mergedStylesXml).toContain('<w:next w:val="Normal"/>');
    } finally {
      doc.dispose();
    }
  });
});
