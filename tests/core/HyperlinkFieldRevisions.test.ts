/**
 * Tests for HYPERLINK field code revisions during parsing and acceptance.
 *
 * Issue: HYPERLINK field codes (fldChar begin/separate/end) with tracked changes
 * (w:ins/w:del) interleaved in the display text were converted to Hyperlink objects,
 * orphaning the revisions as standalone paragraph content. After acceptAllRevisions(),
 * insertion text leaked as plain unlinked text next to the hyperlink.
 *
 * Fix: When revisions exist in the HYPERLINK field result region, keep as ComplexField
 * (which preserves revisions via resultRevisions) instead of converting to Hyperlink.
 */

import { Document } from '../../src/core/Document';
import { ComplexField } from '../../src/elements/Field';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { Revision } from '../../src/elements/Revision';
import { Run } from '../../src/elements/Run';
import { ZipHandler } from '../../src/zip/ZipHandler';

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

/**
 * XML for a HYPERLINK field code with tracked changes in the result region.
 * Structure: fldChar begin → instrText HYPERLINK → fldChar separate → runs + w:ins + w:del → fldChar end
 */
const HYPERLINK_FIELD_WITH_REVISIONS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com/page" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>Co</w:t></w:r>
      <w:ins w:id="101" w:author="Wright, Joseph" w:date="2024-06-15T10:00:00Z">
        <w:r><w:t>mmercial PA Appeals - Co</w:t></w:r>
      </w:ins>
      <w:r><w:t>verage Determination Denial Reasons </w:t></w:r>
      <w:del w:id="102" w:author="Wright, Joseph" w:date="2024-06-15T10:00:00Z">
        <w:r><w:delText>- Appeals Commercial </w:delText></w:r>
      </w:del>
      <w:r><w:t>(CMS-PRD1-086897)</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>`;

/**
 * XML for a normal HYPERLINK field code without revisions.
 */
const HYPERLINK_FIELD_NO_REVISIONS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com/simple" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>Simple Link Text</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>`;

/**
 * XML for a complete HYPERLINK field code entirely inside a w:ins revision.
 * This is Scenario A: the entire field (begin, instrText, separate, result, end) is tracked.
 */
const HYPERLINK_FIELD_INSIDE_REVISION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:ins w:id="300" w:author="Wright, Joseph" w:date="2024-06-15T10:00:00Z">
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com/inserted" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>Inserted Link</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:ins>
    </w:p>
  </w:body>
</w:document>`;

describe('HYPERLINK Field Code Revisions', () => {
  describe('Round-trip preservation', () => {
    it('should keep HYPERLINK field with revisions as ComplexField, not Hyperlink', async () => {
      const buffer = await createDocxBuffer(HYPERLINK_FIELD_WITH_REVISIONS);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const paras = doc.getParagraphs();
      expect(paras.length).toBeGreaterThanOrEqual(1);

      const content = paras[0]!.getContent();

      // Should NOT have been converted to a Hyperlink
      const hyperlinks = content.filter((item) => item instanceof Hyperlink);
      expect(hyperlinks.length).toBe(0);

      // Should be a ComplexField with resultRevisions
      const fields = content.filter((item) => item instanceof ComplexField);
      expect(fields.length).toBe(1);

      const field = fields[0] as ComplexField;
      expect(field.isHyperlinkField()).toBe(true);
      expect(field.hasResultRevisions()).toBe(true);
      expect(field.getResultRevisions().length).toBe(2);

      // Verify revision types
      const revTypes = field.getResultRevisions().map((r) => r.getType());
      expect(revTypes).toContain('insert');
      expect(revTypes).toContain('delete');

      // Verify accepted text includes all non-deleted content
      // Original XML: "Co" + ins("mmercial PA Appeals - Co") + "verage Determination Denial Reasons " + del("- Appeals Commercial ") + "(CMS-PRD1-086897)"
      // Accepted text = "Co" + "mmercial PA Appeals - Co" + "verage Determination Denial Reasons " + "(CMS-PRD1-086897)"
      expect(field.getResult()).toBe(
        'Commercial PA Appeals - Coverage Determination Denial Reasons (CMS-PRD1-086897)'
      );

      // No orphaned revisions at the paragraph level
      const orphanedRevisions = content.filter((item) => item instanceof Revision);
      expect(orphanedRevisions.length).toBe(0);

      doc.dispose();
    });

    it('should preserve revision markup in saved output', async () => {
      const buffer = await createDocxBuffer(HYPERLINK_FIELD_WITH_REVISIONS);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const outputZip = new ZipHandler();
      await outputZip.loadFromBuffer(outputBuffer);
      const outputXml = outputZip.getFileAsString('word/document.xml') || '';

      // Revisions must be inside the field (between separate and end markers)
      expect(outputXml).toContain('w:ins');
      expect(outputXml).toContain('w:del');
      expect(outputXml).toContain('mmercial PA Appeals - Co');
      expect(outputXml).toContain('- Appeals Commercial');
      expect(outputXml).toContain('fldCharType="begin"');
      expect(outputXml).toContain('fldCharType="end"');
    });
  });

  describe('setResult clears revisions', () => {
    it('should clear resultRevisions when setResult is called', async () => {
      const buffer = await createDocxBuffer(HYPERLINK_FIELD_WITH_REVISIONS);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const paras = doc.getParagraphs();
      const content = paras[0]!.getContent();
      const field = content.find((item) => item instanceof ComplexField) as ComplexField;

      expect(field).toBeDefined();
      expect(field.hasResultRevisions()).toBe(true);

      // Update the result text
      field.setResult('New Display Text');

      // Revisions should be cleared
      expect(field.hasResultRevisions()).toBe(false);
      expect(field.getResultRevisions().length).toBe(0);
      expect(field.getResult()).toBe('New Display Text');

      // Save and verify clean output
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const outputZip = new ZipHandler();
      await outputZip.loadFromBuffer(outputBuffer);
      const outputXml = outputZip.getFileAsString('word/document.xml') || '';

      // Old revision text should not appear
      expect(outputXml).not.toContain('mmercial PA Appeals - Co');
      expect(outputXml).not.toContain('- Appeals Commercial');
      // New text should appear
      expect(outputXml).toContain('New Display Text');
    });
  });

  describe('acceptAllRevisions handles field revisions', () => {
    it('should clear resultRevisions when accepting all revisions', async () => {
      const buffer = await createDocxBuffer(HYPERLINK_FIELD_WITH_REVISIONS);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Verify revisions exist before acceptance
      const parasBefore = doc.getParagraphs();
      const contentBefore = parasBefore[0]!.getContent();
      const fieldBefore = contentBefore.find(
        (item) => item instanceof ComplexField
      ) as ComplexField;
      expect(fieldBefore.hasResultRevisions()).toBe(true);

      // Accept all revisions
      doc.acceptAllRevisions();

      // Field should still exist but revisions should be cleared
      const parasAfter = doc.getParagraphs();
      const contentAfter = parasAfter[0]!.getContent();
      const fieldAfter = contentAfter.find((item) => item instanceof ComplexField) as ComplexField;
      expect(fieldAfter).toBeDefined();
      expect(fieldAfter.hasResultRevisions()).toBe(false);

      // After acceptance, result should be the full accepted text in correct interleaved order
      expect(fieldAfter.getResult()).toBe(
        'Commercial PA Appeals - Coverage Determination Denial Reasons (CMS-PRD1-086897)'
      );

      // No orphaned revisions at paragraph level
      const orphanedRevisions = contentAfter.filter((item) => item instanceof Revision);
      expect(orphanedRevisions.length).toBe(0);

      doc.dispose();
    });

    it('should produce correct text when accepting revisions without prior setResult', async () => {
      // This tests the scenario where the app does NOT modify the field (no setResult call)
      // but autoAcceptRevisions is true → acceptAllRevisions() merges insertion text
      const buffer = await createDocxBuffer(HYPERLINK_FIELD_WITH_REVISIONS);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Do NOT call setResult — simulate unmodified field
      doc.acceptAllRevisions();

      const paras = doc.getParagraphs();
      const content = paras[0]!.getContent();
      const field = content.find((item) => item instanceof ComplexField) as ComplexField;

      expect(field).toBeDefined();
      expect(field.hasResultRevisions()).toBe(false);

      // Result should be the full accepted text in correct interleaved order
      expect(field.getResult()).toBe(
        'Commercial PA Appeals - Coverage Determination Denial Reasons (CMS-PRD1-086897)'
      );

      doc.dispose();
    });
  });

  describe('Normal HYPERLINK fields still convert to Hyperlink (regression)', () => {
    it('should convert HYPERLINK field without revisions to Hyperlink object', async () => {
      const buffer = await createDocxBuffer(HYPERLINK_FIELD_NO_REVISIONS);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const paras = doc.getParagraphs();
      expect(paras.length).toBeGreaterThanOrEqual(1);

      const content = paras[0]!.getContent();

      // Should be converted to a Hyperlink, not kept as ComplexField
      const hyperlinks = content.filter((item) => item instanceof Hyperlink);
      expect(hyperlinks.length).toBe(1);

      const complexFields = content.filter((item) => item instanceof ComplexField);
      expect(complexFields.length).toBe(0);

      doc.dispose();
    });
  });

  describe('Interleaved revision ordering in saved XML', () => {
    it('should preserve interleaved revision ordering in saved XML', async () => {
      // Load with preserve → save → verify w:ins appears between "Co" and "verage" in output
      const buffer = await createDocxBuffer(HYPERLINK_FIELD_WITH_REVISIONS);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const outputZip = new ZipHandler();
      await outputZip.loadFromBuffer(outputBuffer);
      const outputXml = outputZip.getFileAsString('word/document.xml') || '';

      // Verify the interleaved order: "Co" run → w:ins → "verage" run → w:del → "(CMS-PRD1-086897)" run
      // Note: use '<w:ins ' (with space) to avoid matching '<w:instrText'
      const coPos = outputXml.indexOf('>Co<');
      const insPos = outputXml.indexOf('<w:ins ');
      const veragePos = outputXml.indexOf('verage Determination');
      const delPos = outputXml.indexOf('<w:del ');
      const cmsPos = outputXml.indexOf('(CMS-PRD1-086897)');

      expect(coPos).toBeGreaterThan(-1);
      expect(insPos).toBeGreaterThan(-1);
      expect(veragePos).toBeGreaterThan(-1);
      expect(delPos).toBeGreaterThan(-1);
      expect(cmsPos).toBeGreaterThan(-1);

      // Assert interleaved order is preserved
      expect(coPos).toBeLessThan(insPos);
      expect(insPos).toBeLessThan(veragePos);
      expect(veragePos).toBeLessThan(delPos);
      expect(delPos).toBeLessThan(cmsPos);
    });
  });

  describe('Entire HYPERLINK field inside revision (Scenario A)', () => {
    it('should keep field inside w:ins as a Revision, not promote to ComplexField', async () => {
      const buffer = await createDocxBuffer(HYPERLINK_FIELD_INSIDE_REVISION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const paras = doc.getParagraphs();
      expect(paras.length).toBeGreaterThanOrEqual(1);

      const content = paras[0]!.getContent();

      // Field should NOT be promoted — it stays as a Revision
      const fields = content.filter((item) => item instanceof ComplexField);
      expect(fields.length).toBe(0);

      const revisions = content.filter((item) => item instanceof Revision);
      expect(revisions.length).toBe(1);

      const revision = revisions[0] as Revision;
      expect(revision.getType()).toBe('insert');

      // Revision should contain Runs with field tokens
      const runs = revision.getContent().filter((c): c is Run => c instanceof Run);
      expect(runs.length).toBeGreaterThanOrEqual(3);

      doc.dispose();
    });

    it('should round-trip w:ins-wrapped field correctly', async () => {
      const buffer = await createDocxBuffer(HYPERLINK_FIELD_INSIDE_REVISION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const outputZip = new ZipHandler();
      await outputZip.loadFromBuffer(outputBuffer);
      const outputXml = outputZip.getFileAsString('word/document.xml') || '';

      // w:ins wrapper must be preserved
      expect(outputXml).toContain('<w:ins ');
      expect(outputXml).toContain('w:id="300"');
      // Field structure inside w:ins
      expect(outputXml).toContain('fldCharType="begin"');
      expect(outputXml).toContain('HYPERLINK');
      expect(outputXml).toContain('https://example.com/inserted');
      expect(outputXml).toContain('Inserted Link');
      expect(outputXml).toContain('fldCharType="end"');
    });

    it('should round-trip w:del-wrapped field with w:delInstrText and w:delText', async () => {
      const deletedFieldXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:del w:id="400" w:author="Test Author" w:date="2024-06-15T10:00:00Z">
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:delInstrText xml:space="preserve"> HYPERLINK "https://example.com/deleted" </w:delInstrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:delText>Deleted Link</w:delText></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:del>
    </w:p>
  </w:body>
</w:document>`;

      const buffer = await createDocxBuffer(deletedFieldXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const paras = doc.getParagraphs();
      const content = paras[0]!.getContent();

      // Should be a Revision, not a ComplexField
      const fields = content.filter((item) => item instanceof ComplexField);
      expect(fields.length).toBe(0);

      const revisions = content.filter((item) => item instanceof Revision);
      expect(revisions.length).toBe(1);
      expect((revisions[0] as Revision).getType()).toBe('delete');

      // Round-trip: save and verify w:del wrapper with w:delInstrText/w:delText
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const outputZip = new ZipHandler();
      await outputZip.loadFromBuffer(outputBuffer);
      const outputXml = outputZip.getFileAsString('word/document.xml') || '';

      expect(outputXml).toContain('<w:del ');
      expect(outputXml).toContain('w:delInstrText');
      expect(outputXml).toContain('w:delText');
      expect(outputXml).toContain('fldCharType="begin"');
      expect(outputXml).toContain('fldCharType="end"');

      // Bare w:instrText should NOT appear — only w:delInstrText inside the w:del wrapper
      expect(outputXml).not.toMatch(/<w:instrText[\s>]/);
    });

    it('should preserve mixed content inside w:ins (text + field)', async () => {
      const mixedContentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:ins w:id="500" w:author="Test Author" w:date="2024-06-15T10:00:00Z">
        <w:r><w:t>See also: </w:t></w:r>
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com/mixed" </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>Mixed Link</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:ins>
    </w:p>
  </w:body>
</w:document>`;

      const buffer = await createDocxBuffer(mixedContentXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const outputZip = new ZipHandler();
      await outputZip.loadFromBuffer(outputBuffer);
      const outputXml = outputZip.getFileAsString('word/document.xml') || '';

      // Both text and field should be preserved inside w:ins
      expect(outputXml).toContain('<w:ins ');
      expect(outputXml).toContain('See also: ');
      expect(outputXml).toContain('fldCharType="begin"');
      expect(outputXml).toContain('HYPERLINK');
      expect(outputXml).toContain('Mixed Link');
      expect(outputXml).toContain('fldCharType="end"');
    });

    it('should remove w:del-wrapped field when accepting revisions', async () => {
      const deletedFieldXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t>Before </w:t></w:r>
      <w:del w:id="400" w:author="Test Author" w:date="2024-06-15T10:00:00Z">
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:delInstrText xml:space="preserve"> HYPERLINK "https://example.com/deleted" </w:delInstrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:delText>Deleted Link</w:delText></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:del>
      <w:r><w:t> After</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

      const buffer = await createDocxBuffer(deletedFieldXml);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      doc.acceptAllRevisions();

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const outputZip = new ZipHandler();
      await outputZip.loadFromBuffer(outputBuffer);
      const outputXml = outputZip.getFileAsString('word/document.xml') || '';

      // Deleted field should be removed entirely
      expect(outputXml).not.toContain('w:del');
      expect(outputXml).not.toContain('Deleted Link');
      expect(outputXml).not.toContain('HYPERLINK');
      // Surrounding content preserved
      expect(outputXml).toContain('Before ');
      expect(outputXml).toContain(' After');
    });

    it('should unwrap w:ins-wrapped field when accepting revisions', async () => {
      const buffer = await createDocxBuffer(HYPERLINK_FIELD_INSIDE_REVISION);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      doc.acceptAllRevisions();

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const outputZip = new ZipHandler();
      await outputZip.loadFromBuffer(outputBuffer);
      const outputXml = outputZip.getFileAsString('word/document.xml') || '';

      // w:ins wrapper should be removed after acceptance
      expect(outputXml).not.toContain('<w:ins ');
      // But the field content should be unwrapped and present
      expect(outputXml).toContain('Inserted Link');
    });
  });
});
