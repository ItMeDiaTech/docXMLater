/**
 * Style-level `<w:ind>` — CJK character-unit indentation round-trip.
 *
 * Iterations 21 & 22 added `w:leftChars` / `w:rightChars` /
 * `w:firstLineChars` / `w:hangingChars` (ECMA-376 §17.3.1.12) support to
 * the main-path paragraph `<w:ind>` AND the pPrChange tracked-change
 * side. BUT the style-level path is a third independent code unit —
 * `DocumentParser.parseParagraphFormattingFromXml` at ~line 9057 parses
 * a style's `<w:pPr><w:ind>` from raw XML, and `Style.toXML()` at
 * ~line 977 emits it from the same shape. Both had the same four-
 * attribute twips-only gap.
 *
 * So even after iterations 21-22, a style definition that carried CJK
 * character-unit indentation round-tripped with those attributes
 * silently dropped. Every paragraph that referenced that style would
 * then render its indent from the twips fallback (or nothing) instead
 * of the source author's character-unit specification.
 */

import { Document } from '../../src/core/Document';
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
  <w:style w:type="paragraph" w:styleId="CJKTest">
    <w:name w:val="CJKTest"/>
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

describe('Style w:ind — CJK character-unit indentation parsing', () => {
  it('parses w:firstLineChars on a style', async () => {
    const buffer = await makeDocxWithStylePPr('<w:ind w:firstLineChars="200"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const style = doc.getStylesManager().getStyle('CJKTest');
    const ind = style?.getParagraphFormatting()?.indentation;
    expect(ind?.firstLineChars).toBe(200);
    doc.dispose();
  });

  it('parses w:hangingChars on a style', async () => {
    const buffer = await makeDocxWithStylePPr('<w:ind w:hangingChars="150"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const style = doc.getStylesManager().getStyle('CJKTest');
    const ind = style?.getParagraphFormatting()?.indentation;
    expect(ind?.hangingChars).toBe(150);
    doc.dispose();
  });

  it('parses w:leftChars on a style', async () => {
    const buffer = await makeDocxWithStylePPr('<w:ind w:leftChars="400"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const style = doc.getStylesManager().getStyle('CJKTest');
    const ind = style?.getParagraphFormatting()?.indentation;
    expect(ind?.leftChars).toBe(400);
    doc.dispose();
  });

  it('parses w:rightChars on a style', async () => {
    const buffer = await makeDocxWithStylePPr('<w:ind w:rightChars="300"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const style = doc.getStylesManager().getStyle('CJKTest');
    const ind = style?.getParagraphFormatting()?.indentation;
    expect(ind?.rightChars).toBe(300);
    doc.dispose();
  });

  it('parses w:startChars (bidi-aware) as leftChars on a style', async () => {
    const buffer = await makeDocxWithStylePPr('<w:ind w:startChars="400"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const style = doc.getStylesManager().getStyle('CJKTest');
    const ind = style?.getParagraphFormatting()?.indentation;
    expect(ind?.leftChars).toBe(400);
    doc.dispose();
  });

  it('parses w:endChars (bidi-aware) as rightChars on a style', async () => {
    const buffer = await makeDocxWithStylePPr('<w:ind w:endChars="300"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const style = doc.getStylesManager().getStyle('CJKTest');
    const ind = style?.getParagraphFormatting()?.indentation;
    expect(ind?.rightChars).toBe(300);
    doc.dispose();
  });

  it('parses combined twips + character-unit indentation on a style', async () => {
    const buffer = await makeDocxWithStylePPr(
      '<w:ind w:left="720" w:leftChars="200" w:firstLine="360" w:firstLineChars="100"/>'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const style = doc.getStylesManager().getStyle('CJKTest');
    const ind = style?.getParagraphFormatting()?.indentation;
    expect(ind?.left).toBe(720);
    expect(ind?.leftChars).toBe(200);
    expect(ind?.firstLine).toBe(360);
    expect(ind?.firstLineChars).toBe(100);
    doc.dispose();
  });
});

describe('Style w:ind — CJK character-unit indentation full round-trip', () => {
  it('round-trips w:firstLineChars on a style through load → save → load', async () => {
    const buffer1 = await makeDocxWithStylePPr('<w:ind w:firstLineChars="200"/>');
    const doc1 = await Document.loadFromBuffer(buffer1);
    expect(
      doc1.getStylesManager().getStyle('CJKTest')?.getParagraphFormatting()?.indentation
        ?.firstLineChars
    ).toBe(200);
    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    expect(
      doc2.getStylesManager().getStyle('CJKTest')?.getParagraphFormatting()?.indentation
        ?.firstLineChars
    ).toBe(200);
    doc2.dispose();
  });

  it('round-trips left/right + firstLine chars on a style', async () => {
    const buffer1 = await makeDocxWithStylePPr(
      '<w:ind w:leftChars="400" w:rightChars="200" w:firstLineChars="100"/>'
    );
    const doc1 = await Document.loadFromBuffer(buffer1);
    const ind1 = doc1.getStylesManager().getStyle('CJKTest')?.getParagraphFormatting()?.indentation;
    expect(ind1?.leftChars).toBe(400);
    expect(ind1?.rightChars).toBe(200);
    expect(ind1?.firstLineChars).toBe(100);

    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    const ind2 = doc2.getStylesManager().getStyle('CJKTest')?.getParagraphFormatting()?.indentation;
    expect(ind2?.leftChars).toBe(400);
    expect(ind2?.rightChars).toBe(200);
    expect(ind2?.firstLineChars).toBe(100);
    doc2.dispose();
  });

  it('round-trips hangingChars as the alternative to firstLineChars on a style', async () => {
    const buffer1 = await makeDocxWithStylePPr('<w:ind w:leftChars="400" w:hangingChars="50"/>');
    const doc1 = await Document.loadFromBuffer(buffer1);
    const ind1 = doc1.getStylesManager().getStyle('CJKTest')?.getParagraphFormatting()?.indentation;
    expect(ind1?.leftChars).toBe(400);
    expect(ind1?.hangingChars).toBe(50);

    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    const ind2 = doc2.getStylesManager().getStyle('CJKTest')?.getParagraphFormatting()?.indentation;
    expect(ind2?.leftChars).toBe(400);
    expect(ind2?.hangingChars).toBe(50);
    doc2.dispose();
  });
});
