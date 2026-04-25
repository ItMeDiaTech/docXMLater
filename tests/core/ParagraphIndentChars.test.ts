/**
 * Paragraph `<w:ind>` — character-based indentation attributes.
 *
 * Per ECMA-376 Part 1 §17.3.1.12 (CT_Ind), `<w:ind>` carries twelve
 * attributes. docxmlater supports the four twips variants
 * (w:left / w:start / w:right / w:end / w:firstLine / w:hanging) but
 * had no support at all for the six character-unit variants:
 *
 *   w:startChars / w:leftChars  — CJK character-unit left indent
 *   w:endChars   / w:rightChars — CJK character-unit right indent
 *   w:firstLineChars            — CJK character-unit first-line indent
 *   w:hangingChars              — CJK character-unit hanging indent
 *
 * Word uses these when the document is authored in an Asian language
 * and the user specifies indentation in "number of character widths"
 * rather than twips. The value is ST_DecimalNumber: Word interprets
 * it as hundredths of a character unit (so "200" means 2 character
 * widths).
 *
 * Bug this suite guards against:
 *   - Grep of entire src/ finds zero references to any of the six
 *     Chars attributes. Parse silently dropped them on load; generator
 *     could never emit them. Any CJK-authored document that used
 *     character-unit indentation (extremely common in real Chinese/
 *     Japanese/Korean Word files) round-tripped with the CJK-specific
 *     indent spec replaced by empty / Word-default twips.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithInd(indXml: string): Promise<Buffer> {
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
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>${indXml}</w:pPr>
      <w:r><w:t>test</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('Paragraph w:ind — CJK character-based indentation parsing', () => {
  it('parses w:firstLineChars', async () => {
    const buffer = await makeDocxWithInd('<w:ind w:firstLineChars="200"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const ind = doc.getParagraphs()[0]!.getFormatting().indentation;
    expect(ind?.firstLineChars).toBe(200);
    doc.dispose();
  });

  it('parses w:hangingChars', async () => {
    const buffer = await makeDocxWithInd('<w:ind w:hangingChars="150"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const ind = doc.getParagraphs()[0]!.getFormatting().indentation;
    expect(ind?.hangingChars).toBe(150);
    doc.dispose();
  });

  it('parses w:leftChars (legacy alias)', async () => {
    const buffer = await makeDocxWithInd('<w:ind w:leftChars="400"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const ind = doc.getParagraphs()[0]!.getFormatting().indentation;
    expect(ind?.leftChars).toBe(400);
    doc.dispose();
  });

  it('parses w:startChars (bidi-aware, preferred) — collapses to leftChars', async () => {
    const buffer = await makeDocxWithInd('<w:ind w:startChars="400"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const ind = doc.getParagraphs()[0]!.getFormatting().indentation;
    // The existing twips parser collapses w:start → leftIndent; character variants
    // follow the same collapse so that downstream consumers see one canonical field.
    expect(ind?.leftChars).toBe(400);
    doc.dispose();
  });

  it('parses w:rightChars', async () => {
    const buffer = await makeDocxWithInd('<w:ind w:rightChars="300"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const ind = doc.getParagraphs()[0]!.getFormatting().indentation;
    expect(ind?.rightChars).toBe(300);
    doc.dispose();
  });

  it('parses w:endChars (bidi-aware) — collapses to rightChars', async () => {
    const buffer = await makeDocxWithInd('<w:ind w:endChars="300"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const ind = doc.getParagraphs()[0]!.getFormatting().indentation;
    expect(ind?.rightChars).toBe(300);
    doc.dispose();
  });

  it('parses firstLineChars=0 (valid zero) without truthy-check drop', async () => {
    // Protect against the same number-0 trap we've been chasing across runs;
    // ST_DecimalNumber allows 0 even if it's an odd value in practice.
    const buffer = await makeDocxWithInd('<w:ind w:firstLineChars="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const ind = doc.getParagraphs()[0]!.getFormatting().indentation;
    expect(ind?.firstLineChars).toBe(0);
    doc.dispose();
  });

  it('parses combined twips + character-unit indentation', async () => {
    const buffer = await makeDocxWithInd(
      '<w:ind w:left="720" w:leftChars="200" w:firstLine="360" w:firstLineChars="100"/>'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const ind = doc.getParagraphs()[0]!.getFormatting().indentation;
    expect(ind?.left).toBe(720);
    expect(ind?.leftChars).toBe(200);
    expect(ind?.firstLine).toBe(360);
    expect(ind?.firstLineChars).toBe(100);
    doc.dispose();
  });
});

describe('Paragraph w:ind — CJK character-based indentation round-trip', () => {
  it('round-trips w:firstLineChars through load → save → load', async () => {
    const buffer1 = await makeDocxWithInd('<w:ind w:firstLineChars="200"/>');
    const doc1 = await Document.loadFromBuffer(buffer1);
    expect(doc1.getParagraphs()[0]!.getFormatting().indentation?.firstLineChars).toBe(200);
    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    expect(doc2.getParagraphs()[0]!.getFormatting().indentation?.firstLineChars).toBe(200);
    doc2.dispose();
  });

  it('round-trips left/right/firstLineChars (realistic CJK indent, hanging mutually exclusive)', async () => {
    // Per ECMA-376 §17.3.1.12, firstLineChars and hangingChars are mutually
    // exclusive — Word never emits both. This test exercises the firstLine
    // branch; see the next test for the hangingChars branch.
    const buffer1 = await makeDocxWithInd(
      '<w:ind w:leftChars="400" w:rightChars="200" w:firstLineChars="100"/>'
    );
    const doc1 = await Document.loadFromBuffer(buffer1);
    const ind1 = doc1.getParagraphs()[0]!.getFormatting().indentation;
    expect(ind1?.leftChars).toBe(400);
    expect(ind1?.rightChars).toBe(200);
    expect(ind1?.firstLineChars).toBe(100);

    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    const ind2 = doc2.getParagraphs()[0]!.getFormatting().indentation;
    expect(ind2?.leftChars).toBe(400);
    expect(ind2?.rightChars).toBe(200);
    expect(ind2?.firstLineChars).toBe(100);
    doc2.dispose();
  });

  it('round-trips hangingChars as the alternative to firstLineChars', async () => {
    const buffer1 = await makeDocxWithInd('<w:ind w:leftChars="400" w:hangingChars="50"/>');
    const doc1 = await Document.loadFromBuffer(buffer1);
    const ind1 = doc1.getParagraphs()[0]!.getFormatting().indentation;
    expect(ind1?.leftChars).toBe(400);
    expect(ind1?.hangingChars).toBe(50);

    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    const ind2 = doc2.getParagraphs()[0]!.getFormatting().indentation;
    expect(ind2?.leftChars).toBe(400);
    expect(ind2?.hangingChars).toBe(50);
    doc2.dispose();
  });
});
