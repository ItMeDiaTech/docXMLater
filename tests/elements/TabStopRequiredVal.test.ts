/**
 * `<w:tab>` inside `<w:tabs>` — required `w:val` attribute compliance.
 *
 * Per ECMA-376 Part 1 §17.3.1.37 CT_TabStop, BOTH `w:val` (ST_TabJc) and
 * `w:pos` (ST_SignedTwipsMeasure) are REQUIRED attributes. Strict OOXML
 * validation rejects any tab emission missing `w:val` with
 * *"The required attribute 'val' is missing"*.
 *
 * Three emission paths had the same truthy-gate `if (tab.val) attrs[...]`
 * that silently dropped `w:val` when callers constructed a tab with only
 * `position` set (a common case — "I want a tab stop at 2 inches, use
 * whatever alignment is default"):
 *
 *   1. `Paragraph.generateParagraphPropertiesXml` — direct paragraph tabs
 *   2. `Paragraph` — pPrChange/previousProperties tabs
 *   3. `Style.generateParagraphProperties` — style-level tabs
 *
 * All three now default `w:val` to `"left"` (matches Word's authored default
 * when no explicit alignment is specified and matches the iteration-58 CT_PTab
 * fix which defaults `w:alignment` to `"left"`).
 */

import { Paragraph } from '../../src/elements/Paragraph';
import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('CT_TabStop w:val required — all emission sites (§17.3.1.37)', () => {
  it('direct paragraph tab defaults w:val to "left" when only position is set', () => {
    const p = new Paragraph();
    p.addText('x');
    p.setTabs([{ position: 720 }]);
    const xml = XMLBuilder.elementToString(p.toXML());
    expect(xml).toMatch(
      /<w:tabs>[\s\S]*?<w:tab[^>]*w:val="left"[^>]*w:pos="720"[\s\S]*?<\/w:tabs>/
    );
    // Must not emit <w:tab w:pos="720"/> without w:val.
    expect(xml).not.toMatch(/<w:tab\s+w:pos="720"\s*\/>/);
  });

  it('direct paragraph tab passes OOXML validator with only position', async () => {
    const doc = Document.create();
    const p = new Paragraph();
    p.addText('tab-test');
    p.setTabs([{ position: 1440 }, { position: 2880, leader: 'dot' }]);
    doc.addParagraph(p);
    await expect(doc.toBuffer()).resolves.toBeInstanceOf(Buffer);
    doc.dispose();
  });

  it('style-level paragraph tab defaults w:val to "left" when only position is set', () => {
    const style = Style.create({
      styleId: 'TabTest',
      name: 'TabTest',
      type: 'paragraph',
    });
    style.setParagraphFormatting({
      tabs: [{ position: 2160 }],
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toMatch(
      /<w:tabs>[\s\S]*?<w:tab[^>]*w:val="left"[^>]*w:pos="2160"[\s\S]*?<\/w:tabs>/
    );
    expect(xml).not.toMatch(/<w:tab\s+w:pos="2160"\s*\/>/);
  });

  it('pPrChange previous tabs default w:val to "left" via validator round-trip', async () => {
    // Author a paragraph whose pPrChange previousProperties includes tabs
    // without w:val. Round-trip through validator to confirm emission fills
    // the required attribute.
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
    <w:p>
      <w:pPr>
        <w:tabs><w:tab w:val="center" w:pos="1440"/></w:tabs>
        <w:pPrChange w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">
          <w:pPr>
            <w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>hello</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
    );
    const buffer = await zipHandler.toBuffer();

    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    // Mutate the first paragraph's pPrChange previous tabs to drop val.
    const para = doc.getParagraphs()[0]!;
    // Access the pPrChange through formatting (no public getter).
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const change = (para as any).formatting.pPrChange;
    if (change?.previousProperties?.tabs) {
      change.previousProperties.tabs = [{ position: 720 }];
    }
    await expect(doc.toBuffer()).resolves.toBeInstanceOf(Buffer);
    doc.dispose();
  });
});
