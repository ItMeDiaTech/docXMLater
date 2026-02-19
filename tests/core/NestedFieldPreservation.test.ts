/**
 * Tests for nested complex field preservation
 * 
 * Verifies that deeply nested fields (e.g., INCLUDEPICTURE multiplication from
 * email-pasted images) are preserved with correct begin/separate/end balance
 * during round-trip through docxmlater.
 * 
 * Bug: assembleComplexFields() had no nesting depth counter, causing nested
 * fields to be flattened into a single ComplexField with concatenated instrTexts
 * and orphaned end markers, corrupting the document's field structure.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

/**
 * Helper function to create a minimal DOCX buffer from document.xml content
 */
async function createDocxFromXml(documentXml: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();

  zipHandler.addFile(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );

  zipHandler.addFile(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  zipHandler.addFile(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`
  );

  zipHandler.addFile("word/document.xml", documentXml);

  return await zipHandler.toBuffer();
}

/**
 * Generates OOXML for N levels of nested INCLUDEPICTURE fields
 * Produces the exact structure Word creates when INCLUDEPICTURE multiplies on save
 */
function generateNestedIncludePicture(depth: number): string {
  let beginRuns = '';
  let endRuns = '';

  for (let i = 0; i < depth; i++) {
    beginRuns += `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`;
    beginRuns += `<w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE  "cid:image004.png@01DB4D78.DF89DCD0" \\* MERGEFORMATINET </w:instrText></w:r>`;
    beginRuns += `<w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
    endRuns += `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  }

  // The innermost result is a simple text run
  const resultRun = `<w:r><w:t>[Image Placeholder]</w:t></w:r>`;

  return beginRuns + resultRun + endRuns;
}

describe('Nested Complex Field Preservation', () => {
  it('should preserve balanced field structure for 3-level nested INCLUDEPICTURE', async () => {
    const nestedField = generateNestedIncludePicture(3);
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            mc:Ignorable="w14">
  <w:body>
    <w:p>
      <w:r><w:t>Before field </w:t></w:r>
      ${nestedField}
      <w:r><w:t> After field</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createDocxFromXml(documentXml);
    const doc = await Document.loadFromBuffer(buffer);
    const outputBuffer = await doc.toBuffer();
    const outputDoc = await Document.loadFromBuffer(outputBuffer);

    // Extract the output XML to verify field balance
    const outputZip = new ZipHandler();
    await outputZip.loadFromBuffer(outputBuffer);
    const outputXml = outputZip.getFileAsString('word/document.xml')!;

    // Count field chars
    const beginCount = (outputXml.match(/w:fldCharType="begin"/g) || []).length;
    const separateCount = (outputXml.match(/w:fldCharType="separate"/g) || []).length;
    const endCount = (outputXml.match(/w:fldCharType="end"/g) || []).length;

    // CRITICAL: begins must equal ends for valid OOXML
    expect(beginCount).toBe(endCount);
    expect(beginCount).toBe(3);
    expect(separateCount).toBe(3);
  });

  it('should preserve balanced field structure for deeply nested INCLUDEPICTURE (10 levels)', async () => {
    const nestedField = generateNestedIncludePicture(10);
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            mc:Ignorable="w14">
  <w:body>
    <w:p>
      ${nestedField}
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createDocxFromXml(documentXml);
    const doc = await Document.loadFromBuffer(buffer);
    const outputBuffer = await doc.toBuffer();

    const outputZip = new ZipHandler();
    await outputZip.loadFromBuffer(outputBuffer);
    const outputXml = outputZip.getFileAsString('word/document.xml')!;

    const beginCount = (outputXml.match(/w:fldCharType="begin"/g) || []).length;
    const endCount = (outputXml.match(/w:fldCharType="end"/g) || []).length;

    expect(beginCount).toBe(endCount);
    expect(beginCount).toBe(10);
  });

  it('should still correctly assemble non-nested (flat) complex fields', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            mc:Ignorable="w14">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>1</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createDocxFromXml(documentXml);
    const doc = await Document.loadFromBuffer(buffer);
    const outputBuffer = await doc.toBuffer();

    const outputZip = new ZipHandler();
    await outputZip.loadFromBuffer(outputBuffer);
    const outputXml = outputZip.getFileAsString('word/document.xml')!;

    const beginCount = (outputXml.match(/w:fldCharType="begin"/g) || []).length;
    const endCount = (outputXml.match(/w:fldCharType="end"/g) || []).length;

    // Non-nested field should still be assembled correctly
    expect(beginCount).toBe(endCount);
    expect(beginCount).toBe(1);
  });

  it('should handle mixed nested and non-nested fields in the same paragraph', async () => {
    const nestedField = generateNestedIncludePicture(3);
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            mc:Ignorable="w14">
  <w:body>
    <w:p>
      ${nestedField}
      <w:r><w:t> Page: </w:t></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>1</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createDocxFromXml(documentXml);
    const doc = await Document.loadFromBuffer(buffer);
    const outputBuffer = await doc.toBuffer();

    const outputZip = new ZipHandler();
    await outputZip.loadFromBuffer(outputBuffer);
    const outputXml = outputZip.getFileAsString('word/document.xml')!;

    const beginCount = (outputXml.match(/w:fldCharType="begin"/g) || []).length;
    const separateCount = (outputXml.match(/w:fldCharType="separate"/g) || []).length;
    const endCount = (outputXml.match(/w:fldCharType="end"/g) || []).length;

    // 3 nested INCLUDEPICTURE + 1 PAGE field = 4 total
    expect(beginCount).toBe(endCount);
    expect(beginCount).toBe(4);
    expect(separateCount).toBe(4);
  });
});
