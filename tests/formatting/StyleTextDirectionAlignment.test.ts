/**
 * Style-level pPr — textDirection, textAlignment, textboxTightWrap.
 *
 * CT_PPrBase positions #28 (textDirection §17.3.1.36), #29 (textAlignment
 * §17.3.1.39 — vertical run alignment), and #30 (textboxTightWrap
 * §17.3.1.37). All three are valid on a paragraph style's pPr.
 *
 * The main paragraph parser handles all three at DocumentParser.ts:~2450-2548.
 * The style-level parser (`parseParagraphFormattingFromXml`) and the style
 * serializer (`Style.generateParagraphProperties`) silently drop them — a
 * paragraph style that sets `textDirection: "tbRl"` to force vertical CJK
 * flow would lose the override on any programmatic save that bypasses raw
 * XML passthrough (e.g. after modifying the style).
 */

import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStylePPr(pPrInner: string): Promise<Buffer> {
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
  <w:style w:type="paragraph" w:styleId="TdTaTtw">
    <w:name w:val="TdTaTtw"/>
    <w:pPr>${pPrInner}</w:pPr>
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
  return doc.getStylesManager().getStyle('TdTaTtw')?.getParagraphFormatting();
}

describe('Style pPr — textDirection (ECMA-376 §17.3.1.36)', () => {
  const DIRS = ['lrTb', 'tbRl', 'btLr', 'lrTbV', 'tbRlV', 'tbLrV'] as const;

  for (const dir of DIRS) {
    it(`parses <w:textDirection w:val="${dir}"/> as "${dir}"`, async () => {
      const buffer = await makeDocxWithStylePPr(`<w:textDirection w:val="${dir}"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const fmt = getFmt(doc);
      expect(fmt?.textDirection).toBe(dir);
      doc.dispose();
    });

    it(`emits <w:textDirection w:val="${dir}"/> via toXML()`, () => {
      const style = new Style({
        styleId: 'TdTaTtw',
        type: 'paragraph',
        name: 'TdTaTtw',
        paragraphFormatting: { textDirection: dir },
      });
      const xml = XMLBuilder.elementToString(style.toXML());
      expect(xml).toContain(`<w:textDirection w:val="${dir}"/>`);
    });
  }

  it('absent textDirection remains undefined', async () => {
    const buffer = await makeDocxWithStylePPr('');
    const doc = await Document.loadFromBuffer(buffer);
    const fmt = getFmt(doc);
    expect(fmt?.textDirection).toBeUndefined();
    doc.dispose();
  });

  it('omits <w:textDirection/> when undefined', () => {
    const style = new Style({
      styleId: 'TdTaTtw',
      type: 'paragraph',
      name: 'TdTaTtw',
      paragraphFormatting: {},
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).not.toContain('<w:textDirection');
  });
});

describe('Style pPr — textAlignment (ECMA-376 §17.3.1.39)', () => {
  const ALIGNS = ['top', 'center', 'baseline', 'bottom', 'auto'] as const;

  for (const align of ALIGNS) {
    it(`parses <w:textAlignment w:val="${align}"/> as "${align}"`, async () => {
      const buffer = await makeDocxWithStylePPr(`<w:textAlignment w:val="${align}"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const fmt = getFmt(doc);
      expect(fmt?.textAlignment).toBe(align);
      doc.dispose();
    });

    it(`emits <w:textAlignment w:val="${align}"/> via toXML()`, () => {
      const style = new Style({
        styleId: 'TdTaTtw',
        type: 'paragraph',
        name: 'TdTaTtw',
        paragraphFormatting: { textAlignment: align },
      });
      const xml = XMLBuilder.elementToString(style.toXML());
      expect(xml).toContain(`<w:textAlignment w:val="${align}"/>`);
    });
  }

  it('absent textAlignment remains undefined', async () => {
    const buffer = await makeDocxWithStylePPr('');
    const doc = await Document.loadFromBuffer(buffer);
    const fmt = getFmt(doc);
    expect(fmt?.textAlignment).toBeUndefined();
    doc.dispose();
  });
});

describe('Style pPr — textboxTightWrap (ECMA-376 §17.3.1.37)', () => {
  const WRAPS = ['none', 'allLines', 'firstAndLastLine', 'firstLineOnly', 'lastLineOnly'] as const;

  for (const wrap of WRAPS) {
    it(`parses <w:textboxTightWrap w:val="${wrap}"/> as "${wrap}"`, async () => {
      const buffer = await makeDocxWithStylePPr(`<w:textboxTightWrap w:val="${wrap}"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const fmt = getFmt(doc);
      expect(fmt?.textboxTightWrap).toBe(wrap);
      doc.dispose();
    });

    it(`emits <w:textboxTightWrap w:val="${wrap}"/> via toXML()`, () => {
      const style = new Style({
        styleId: 'TdTaTtw',
        type: 'paragraph',
        name: 'TdTaTtw',
        paragraphFormatting: { textboxTightWrap: wrap },
      });
      const xml = XMLBuilder.elementToString(style.toXML());
      expect(xml).toContain(`<w:textboxTightWrap w:val="${wrap}"/>`);
    });
  }
});

describe('Style pPr — CT_PPrBase schema ordering (§17.3.1.26)', () => {
  it('emits children in the documented order: jc → textDirection → textAlignment → textboxTightWrap → outlineLvl', () => {
    const style = new Style({
      styleId: 'TdTaTtw',
      type: 'paragraph',
      name: 'TdTaTtw',
      paragraphFormatting: {
        alignment: 'center',
        textDirection: 'tbRl',
        textAlignment: 'top',
        textboxTightWrap: 'allLines',
        outlineLevel: 2,
      },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    const jcIdx = xml.indexOf('<w:jc ');
    const tdIdx = xml.indexOf('<w:textDirection ');
    const taIdx = xml.indexOf('<w:textAlignment ');
    const ttwIdx = xml.indexOf('<w:textboxTightWrap ');
    const olIdx = xml.indexOf('<w:outlineLvl ');
    expect(jcIdx).toBeGreaterThan(-1);
    expect(tdIdx).toBeGreaterThan(jcIdx);
    expect(taIdx).toBeGreaterThan(tdIdx);
    expect(ttwIdx).toBeGreaterThan(taIdx);
    expect(olIdx).toBeGreaterThan(ttwIdx);
  });
});
