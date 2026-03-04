/**
 * Tests for ComplexField.getAcceptedResultText() and setTrackedResult().
 *
 * Validates that field code hyperlink display text can be updated with
 * proper tracked changes (w:del/w:ins pairs) in the resultContent.
 */

import { Document } from '../../src/core/Document';
import { ComplexField } from '../../src/elements/Field';
import { ZipHandler } from '../../src/zip/ZipHandler';
import type { XMLElement } from '../../src/xml/XMLBuilder';

/**
 * Helper to create a minimal DOCX buffer with custom document.xml
 */
async function createDocxBuffer(documentXml: string): Promise<Buffer> {
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
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`
  );

  zipHandler.addFile('word/document.xml', documentXml);

  return await zipHandler.toBuffer();
}

/** Helper to get the first child element by name from an XMLElement */
function findChild(el: XMLElement, name: string): XMLElement | undefined {
  return (el.children || []).find((c): c is XMLElement => typeof c !== 'string' && c.name === name);
}

/** Helper to get text content from first matching child */
function getChildText(el: XMLElement, name: string): string | undefined {
  const child = findChild(el, name);
  if (!child) return undefined;
  const textNode = (child.children || []).find((c): c is string => typeof c === 'string');
  return textNode;
}

// Simple HYPERLINK field with plain text result
const SIMPLE_HYPERLINK_FIELD = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><w:t>Example Link</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>`;

// HYPERLINK field with interleaved tracked changes in result
const FIELD_WITH_REVISIONS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com/doc" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>Original </w:t></w:r>
      <w:ins w:id="201" w:author="Test Author" w:date="2024-06-15T10:00:00Z">
        <w:r><w:t>Inserted </w:t></w:r>
      </w:ins>
      <w:del w:id="202" w:author="Test Author" w:date="2024-06-15T10:00:00Z">
        <w:r><w:delText>Deleted </w:delText></w:r>
      </w:del>
      <w:r><w:t>Text</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>`;

describe('ComplexField.getAcceptedResultText()', () => {
  it('returns result text for simple field (no resultContent)', () => {
    const field = new ComplexField({
      instruction: ' HYPERLINK "https://example.com" ',
      result: 'Simple Text',
    });
    expect(field.getAcceptedResultText()).toBe('Simple Text');
  });

  it('returns empty string for field with no result and no resultContent', () => {
    const field = new ComplexField({
      instruction: ' PAGE ',
    });
    expect(field.getAcceptedResultText()).toBe('');
  });

  it('extracts text from plain runs in resultContent', () => {
    const field = new ComplexField({
      instruction: ' HYPERLINK "https://example.com" ',
      resultContent: [
        {
          name: 'w:r',
          children: [{ name: 'w:t', children: ['Example '] }],
        },
        {
          name: 'w:r',
          children: [{ name: 'w:t', children: ['Link'] }],
        },
      ],
    });

    expect(field.getAcceptedResultText()).toBe('Example Link');
  });

  it('extracts text from interleaved ins/del resultContent (accepted view)', async () => {
    const buf = await createDocxBuffer(FIELD_WITH_REVISIONS);
    const doc = await Document.loadFromBuffer(buf, { revisionHandling: 'preserve' });

    const paragraphs = doc.getAllParagraphs();
    const field = paragraphs[0]!
      .getContent()
      .find((item): item is ComplexField => item instanceof ComplexField);

    expect(field).toBeDefined();
    // Accepted text = plain runs + insertion text, skip deletion
    expect(field!.getAcceptedResultText()).toBe('Original Inserted Text');

    doc.dispose();
  });
});

describe('ComplexField.setTrackedResult()', () => {
  it('creates del/ins pair for simple field', () => {
    const field = new ComplexField({
      instruction: ' HYPERLINK "https://example.com" ',
      result: 'Old Text',
    });

    field.setTrackedResult('New Text', 'Test Author');

    expect(field.getResult()).toBe('New Text');

    const rc = field.getResultContent();
    expect(rc).toHaveLength(2);

    const del = rc[0]!;
    const ins = rc[1]!;
    expect(del.name).toBe('w:del');
    expect(ins.name).toBe('w:ins');

    // del should contain old text
    const delRun = findChild(del, 'w:r')!;
    expect(getChildText(delRun, 'w:delText')).toBe('Old Text');

    // ins should contain new text
    const insRun = findChild(ins, 'w:r')!;
    expect(getChildText(insRun, 'w:t')).toBe('New Text');

    // Author and date should be set
    expect(del.attributes!['w:author']).toBe('Test Author');
    expect(ins.attributes!['w:author']).toBe('Test Author');
    expect(del.attributes!['w:date']).toBeDefined();
    expect(ins.attributes!['w:date']).toBeDefined();
  });

  it('is a no-op when text is unchanged', () => {
    const field = new ComplexField({
      instruction: ' HYPERLINK "https://example.com" ',
      result: 'Same Text',
    });

    field.setTrackedResult('Same Text', 'Author');

    // resultContent should remain empty (no tracked change created)
    expect(field.getResultContent()).toHaveLength(0);
    expect(field.getResult()).toBe('Same Text');
  });

  it('handles field with pre-existing revisions in resultContent', async () => {
    const buf = await createDocxBuffer(FIELD_WITH_REVISIONS);
    const doc = await Document.loadFromBuffer(buf, { revisionHandling: 'preserve' });

    const paragraphs = doc.getAllParagraphs();
    const field = paragraphs[0]!
      .getContent()
      .find((item): item is ComplexField => item instanceof ComplexField);

    expect(field).toBeDefined();

    // Current accepted text: "Original Inserted Text"
    field!.setTrackedResult('Updated Display', 'New Author');

    expect(field!.getResult()).toBe('Updated Display');

    const rc = field!.getResultContent();
    expect(rc).toHaveLength(2);

    const del = rc[0]!;
    const ins = rc[1]!;
    expect(del.name).toBe('w:del');
    expect(ins.name).toBe('w:ins');

    // Del should wrap the accepted text from the old content
    const delRun = findChild(del, 'w:r')!;
    expect(getChildText(delRun, 'w:delText')).toBe('Original Inserted Text');

    // Ins should have the new text
    const insRun = findChild(ins, 'w:r')!;
    expect(getChildText(insRun, 'w:t')).toBe('Updated Display');

    doc.dispose();
  });

  it('preserves formatting from first visible run with preserveFormatting option', () => {
    const field = new ComplexField({
      instruction: ' HYPERLINK "https://example.com" ',
      resultContent: [
        {
          name: 'w:r',
          children: [
            {
              name: 'w:rPr',
              children: [
                { name: 'w:color', attributes: { 'w:val': '0000FF' }, selfClosing: true },
                { name: 'w:u', attributes: { 'w:val': 'single' }, selfClosing: true },
              ],
            },
            { name: 'w:t', children: ['Old Link Text'] },
          ],
        },
      ],
    });

    field.setTrackedResult('New Link Text', 'Author', { preserveFormatting: true });

    const rc = field.getResultContent();
    expect(rc).toHaveLength(2);

    // Both del and ins runs should have rPr copied from the original
    const delRun = findChild(rc[0]!, 'w:r')!;
    const delRPr = findChild(delRun, 'w:rPr');
    expect(delRPr).toBeDefined();

    const insRun = findChild(rc[1]!, 'w:r')!;
    const insRPr = findChild(insRun, 'w:rPr');
    expect(insRPr).toBeDefined();

    // Verify the rPr content matches the original
    expect(findChild(delRPr!, 'w:color')).toBeDefined();
    expect(findChild(insRPr!, 'w:color')).toBeDefined();
  });

  it('applies explicit formatting to both del and ins runs', () => {
    const field = new ComplexField({
      instruction: ' HYPERLINK "https://example.com" ',
      result: 'Old Text',
    });

    field.setTrackedResult('New Text', 'Author', {
      formatting: { color: '0000FF', underline: 'single' },
    });

    const rc = field.getResultContent();

    const delRun = findChild(rc[0]!, 'w:r')!;
    expect(findChild(delRun, 'w:rPr')).toBeDefined();

    const insRun = findChild(rc[1]!, 'w:r')!;
    expect(findChild(insRun, 'w:rPr')).toBeDefined();
  });

  it('clears resultRevisions after setTrackedResult()', () => {
    const field = new ComplexField({
      instruction: ' HYPERLINK "https://example.com" ',
      result: 'Text',
    });

    const mockRevision = { type: 'insert' } as any;
    field.setResultRevisions([mockRevision]);
    expect(field.hasResultRevisions()).toBe(true);

    field.setTrackedResult('New Text', 'Author');

    expect(field.hasResultRevisions()).toBe(false);
  });
});

describe('ComplexField.setTrackedResult() round-trip', () => {
  it('saves document with tracked result changes (del/ins in output XML)', async () => {
    // Use field with revisions — stays as ComplexField after parsing
    const buf = await createDocxBuffer(FIELD_WITH_REVISIONS);
    const doc = await Document.loadFromBuffer(buf, { revisionHandling: 'preserve' });

    const paragraphs = doc.getAllParagraphs();
    const field = paragraphs[0]!
      .getContent()
      .find((item): item is ComplexField => item instanceof ComplexField);
    expect(field).toBeDefined();

    field!.setTrackedResult('Updated Display', 'DocHub');

    // Verify toXML produces correct output before save
    const xml = field!.toXML();
    const delEl = xml.find((el) => el.name === 'w:del');
    const insEl = xml.find((el) => el.name === 'w:ins');
    expect(delEl).toBeDefined();
    expect(insEl).toBeDefined();

    // Verify del contains accepted old text
    const delRun = findChild(delEl!, 'w:r')!;
    expect(getChildText(delRun, 'w:delText')).toBe('Original Inserted Text');

    // Verify ins contains new text
    const insRun = findChild(insEl!, 'w:r')!;
    expect(getChildText(insRun, 'w:t')).toBe('Updated Display');

    // Save should succeed without errors
    const savedBuf = await doc.toBuffer();
    expect(savedBuf.length).toBeGreaterThan(0);

    doc.dispose();
  });

  it('toXML() emits del/ins in result section', () => {
    const field = new ComplexField({
      instruction: ' HYPERLINK "https://example.com" ',
      result: 'Old Link',
    });

    field.setTrackedResult('New Link', 'Author');

    const xml = field.toXML();
    // begin + instr + sep + del + ins + end = 6
    expect(xml.length).toBeGreaterThanOrEqual(6);

    const delEl = xml.find((el) => el.name === 'w:del');
    const insEl = xml.find((el) => el.name === 'w:ins');
    expect(delEl).toBeDefined();
    expect(insEl).toBeDefined();

    const delRun = findChild(delEl!, 'w:r')!;
    expect(getChildText(delRun, 'w:delText')).toBe('Old Link');

    const insRun = findChild(insEl!, 'w:r')!;
    expect(getChildText(insRun, 'w:t')).toBe('New Link');
  });
});

describe('ComplexField tracking context integration', () => {
  it('receives tracking context when document has track changes enabled', async () => {
    // Use field with revisions — stays as ComplexField after parsing
    const buf = await createDocxBuffer(FIELD_WITH_REVISIONS);
    const doc = await Document.loadFromBuffer(buf, { revisionHandling: 'preserve' });

    doc.enableTrackChanges({ author: 'Test Author' });

    const paragraphs = doc.getAllParagraphs();
    const field = paragraphs[0]!
      .getContent()
      .find((item): item is ComplexField => item instanceof ComplexField);
    expect(field).toBeDefined();

    field!.setTrackedResult('Tracked Update', 'Test Author');

    const rc = field!.getResultContent();
    expect(rc).toHaveLength(2);

    const delId = parseInt(rc[0]!.attributes!['w:id'] as string, 10);
    const insId = parseInt(rc[1]!.attributes!['w:id'] as string, 10);
    expect(delId).toBeGreaterThan(0);
    expect(insId).toBeGreaterThan(delId);

    doc.dispose();
  });
});
