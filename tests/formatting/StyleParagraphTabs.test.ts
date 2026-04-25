/**
 * Style-level pPr — tab stops (w:tabs) round-trip.
 *
 * CT_PPrBase position #11 (w:tabs §17.3.1.38 — paragraph tab stops).
 * The style-level parser reads tab stops into `paragraphFormatting.tabs`
 * but `Style.generateParagraphProperties` was silently dropping them on
 * save. Any paragraph style overriding tab stops was lost on programmatic
 * resave that bypassed raw-XML passthrough.
 */

import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyleTabs(tabsXml: string): Promise<Buffer> {
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
  <w:style w:type="paragraph" w:styleId="TabsTest">
    <w:name w:val="TabsTest"/>
    <w:pPr>${tabsXml}</w:pPr>
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
  return doc.getStylesManager().getStyle('TabsTest')?.getParagraphFormatting();
}

describe('Style pPr — tab stops (§17.3.1.38)', () => {
  it('parses <w:tabs> with multiple tab stops', async () => {
    const buffer = await makeDocxWithStyleTabs(
      `<w:tabs>
        <w:tab w:val="left" w:pos="720"/>
        <w:tab w:val="center" w:pos="3600" w:leader="dot"/>
        <w:tab w:val="right" w:pos="9360"/>
      </w:tabs>`
    );
    const doc = await Document.loadFromBuffer(buffer);
    const tabs = getFmt(doc)?.tabs;
    expect(tabs).toBeDefined();
    expect(tabs?.length).toBe(3);
    expect(tabs?.[0]).toMatchObject({ position: 720, val: 'left' });
    expect(tabs?.[1]).toMatchObject({ position: 3600, val: 'center', leader: 'dot' });
    expect(tabs?.[2]).toMatchObject({ position: 9360, val: 'right' });
    doc.dispose();
  });

  it('emits <w:tabs> via Style.toXML() when tabs are set', () => {
    const style = new Style({
      styleId: 'TabsTest',
      type: 'paragraph',
      name: 'TabsTest',
      paragraphFormatting: {
        tabs: [
          { position: 720, val: 'left' },
          { position: 3600, val: 'center', leader: 'dot' },
        ],
      },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toContain('<w:tabs>');
    expect(xml).toContain('w:pos="720"');
    expect(xml).toContain('w:val="left"');
    expect(xml).toContain('w:pos="3600"');
    expect(xml).toContain('w:val="center"');
    expect(xml).toContain('w:leader="dot"');
    expect(xml).toContain('</w:tabs>');
  });

  it('omits <w:tabs/> when tabs array is empty or undefined', () => {
    const styleA = new Style({
      styleId: 'TabsTest',
      type: 'paragraph',
      name: 'TabsTest',
      paragraphFormatting: {},
    });
    const styleB = new Style({
      styleId: 'TabsTest',
      type: 'paragraph',
      name: 'TabsTest',
      paragraphFormatting: { tabs: [] },
    });
    expect(XMLBuilder.elementToString(styleA.toXML())).not.toContain('<w:tabs');
    expect(XMLBuilder.elementToString(styleB.toXML())).not.toContain('<w:tabs');
  });

  it('emits tabs between shd (#10) and suppressAutoHyphens (#12) per CT_PPrBase', () => {
    const style = new Style({
      styleId: 'TabsTest',
      type: 'paragraph',
      name: 'TabsTest',
      paragraphFormatting: {
        shading: { pattern: 'clear', fill: 'FFFF00' },
        tabs: [{ position: 720, val: 'left' }],
        suppressAutoHyphens: true,
      },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    const shdIdx = xml.indexOf('<w:shd');
    const tabsIdx = xml.indexOf('<w:tabs>');
    const sahIdx = xml.indexOf('<w:suppressAutoHyphens');
    expect(shdIdx).toBeGreaterThan(-1);
    expect(tabsIdx).toBeGreaterThan(shdIdx);
    expect(sahIdx).toBeGreaterThan(tabsIdx);
  });
});
