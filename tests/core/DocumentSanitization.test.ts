/**
 * Document Sanitization Tests
 *
 * Tests cover:
 * 1. flattenFieldCodes() — INCLUDEPICTURE field flattening
 * 2. stripOrphanRSIDs() — orphan RSID removal from settings.xml
 * 3. Combined sanitization — chained methods
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

/**
 * Helper: Creates a minimal DOCX buffer with custom document.xml content.
 * Builds a valid DOCX first, then post-processes the ZIP to inject custom XML.
 */
async function createDocxWithDocumentXml(documentXml: string): Promise<Buffer> {
  const doc = Document.create();
  doc.addParagraph(new Paragraph().addText('placeholder'));
  const buffer = await doc.toBuffer();
  doc.dispose();

  const zipHandler = new ZipHandler();
  await zipHandler.loadFromBuffer(buffer);
  zipHandler.updateFile(DOCX_PATHS.DOCUMENT, documentXml);
  return await zipHandler.toBuffer();
}

/**
 * Helper: Creates a DOCX buffer with custom document.xml AND settings.xml.
 */
async function createDocxWithDocAndSettings(documentXml: string, settingsXml: string): Promise<Buffer> {
  const doc = Document.create();
  doc.addParagraph(new Paragraph().addText('placeholder'));
  const buffer = await doc.toBuffer();
  doc.dispose();

  const zipHandler = new ZipHandler();
  await zipHandler.loadFromBuffer(buffer);
  zipHandler.updateFile(DOCX_PATHS.DOCUMENT, documentXml);
  zipHandler.updateFile(DOCX_PATHS.SETTINGS, settingsXml);
  return await zipHandler.toBuffer();
}

// === Minimal document.xml wrapper ===
const DOC_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
            mc:Ignorable="w14 wp14">
  <w:body>`;

const DOC_FOOTER = `    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`;

function wrapDocXml(bodyContent: string): string {
  return `${DOC_HEADER}\n${bodyContent}\n${DOC_FOOTER}`;
}

// === Test XML fragments ===

/** Single INCLUDEPICTURE field wrapping a w:drawing */
const SINGLE_INCLUDEPICTURE = `
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE "cid:image001.png@01DB4D78" \\* MERGEFORMAT </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>`;

/** 3-level nested INCLUDEPICTURE (each level wraps the next) */
const NESTED_INCLUDEPICTURE_3 = `
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE "cid:image001.png" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE "cid:image001.png" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE "cid:image001.png" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>`;

/** HYPERLINK field — should NOT be touched */
const HYPERLINK_FIELD = `
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Click here</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>`;

/** TOC field — should NOT be touched */
const TOC_FIELD = `
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>Table of Contents placeholder</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>`;

/** Plain paragraph with no fields */
const PLAIN_PARAGRAPH = `
    <w:p>
      <w:r><w:t>Hello World</w:t></w:r>
    </w:p>`;

// === Settings XML templates ===

/** Settings with many RSIDs, most orphaned */
function settingsWithRsids(rsidRoot: string, rsidValues: string[]): string {
  const rsidElements = rsidValues.map(v => `    <w:rsid w:val="${v}"/>`).join('\n');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
  <w:rsids>
    <w:rsidRoot w:val="${rsidRoot}"/>
${rsidElements}
  </w:rsids>
  <w:themeFontLang w:val="en-US"/>
</w:settings>`;
}

/** Document XML that references specific RSIDs */
function documentWithRsids(rsidValues: string[]): string {
  const paras = rsidValues.map(v =>
    `    <w:p w:rsidR="${v}" w:rsidRDefault="${v}"><w:r><w:t>Content</w:t></w:r></w:p>`
  ).join('\n');
  return wrapDocXml(paras);
}

// =============================================================================
// Tests
// =============================================================================

describe('Document Sanitization', () => {

  // =========================================================================
  // flattenFieldCodes()
  // =========================================================================
  describe('flattenFieldCodes()', () => {

    it('strips single INCLUDEPICTURE field, preserves w:drawing content', async () => {
      const docXml = wrapDocXml(SINGLE_INCLUDEPICTURE);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes();
      const output = await doc.toBuffer();
      doc.dispose();

      // Verify: no fldChar remains for INCLUDEPICTURE
      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // The drawing content should still be present
      expect(resultXml).toContain('w:drawing');
      expect(resultXml).toContain('r:embed="rId5"');

      // No INCLUDEPICTURE instrText should remain
      expect(resultXml).not.toContain('INCLUDEPICTURE');

      // No fldChar elements should remain (from the INCLUDEPICTURE field)
      expect(resultXml).not.toContain('w:fldChar');
    });

    it('collapses 3-level nested INCLUDEPICTURE to just image content', async () => {
      const docXml = wrapDocXml(NESTED_INCLUDEPICTURE_3);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes();
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Drawing preserved
      expect(resultXml).toContain('w:drawing');
      expect(resultXml).toContain('r:embed="rId5"');

      // All 3 levels of INCLUDEPICTURE stripped
      expect(resultXml).not.toContain('INCLUDEPICTURE');
      expect(resultXml).not.toContain('w:fldChar');
      expect(resultXml).not.toContain('w:instrText');
    });

    it('does NOT touch HYPERLINK or TOC fields', async () => {
      const docXml = wrapDocXml(HYPERLINK_FIELD + TOC_FIELD);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes();
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Both fields should still have their structure intact
      expect(resultXml).toContain('HYPERLINK');
      expect(resultXml).toContain('TOC');
      expect(resultXml).toContain('Click here');
      expect(resultXml).toContain('Table of Contents placeholder');

      // fldChar elements should still be present for these fields
      // Each field has 3 fldChar elements (begin, separate, end) x 2 fields = 6
      const fldCharBeginCount = (resultXml.match(/w:fldCharType\s*=\s*"begin"/g) || []).length;
      const fldCharEndCount = (resultXml.match(/w:fldCharType\s*=\s*"end"/g) || []).length;
      expect(fldCharBeginCount).toBe(2);
      expect(fldCharEndCount).toBe(2);
    });

    it('is a no-op on document with no INCLUDEPICTURE fields', async () => {
      const docXml = wrapDocXml(PLAIN_PARAGRAPH + HYPERLINK_FIELD);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes();
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Plain text preserved
      expect(resultXml).toContain('Hello World');
      // HYPERLINK field preserved
      expect(resultXml).toContain('HYPERLINK');
      expect(resultXml).toContain('Click here');
    });

    it('handles mixed INCLUDEPICTURE and other fields in same document', async () => {
      const docXml = wrapDocXml(HYPERLINK_FIELD + SINGLE_INCLUDEPICTURE + TOC_FIELD);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes();
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // INCLUDEPICTURE removed, drawing kept
      expect(resultXml).not.toContain('INCLUDEPICTURE');
      expect(resultXml).toContain('w:drawing');

      // Other fields preserved
      expect(resultXml).toContain('HYPERLINK');
      expect(resultXml).toContain('TOC');
    });
  });

  // =========================================================================
  // stripOrphanRSIDs()
  // =========================================================================
  describe('stripOrphanRSIDs()', () => {

    it('removes RSIDs not referenced in document.xml', async () => {
      const referencedRsids = ['00AA1111', '00BB2222'];
      const orphanRsids = ['00CC3333', '00DD4444', '00EE5555', '00FF6666'];
      const allRsids = [...referencedRsids, ...orphanRsids];
      const rsidRoot = '00AA1111';

      const docXml = documentWithRsids(referencedRsids);
      const setXml = settingsWithRsids(rsidRoot, allRsids);
      const buffer = await createDocxWithDocAndSettings(docXml, setXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Verify all RSIDs loaded
      expect(doc.getRsids().length).toBe(allRsids.length);

      // preserveRawXml keeps the original document.xml (with rsid attributes)
      // so the RSID scan can find which ones are actually referenced
      doc.preserveRawXml();
      doc.stripOrphanRSIDs();
      const output = await doc.toBuffer();
      doc.dispose();

      // Check output settings.xml
      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const settingsOut = zip.getFileAsString(DOCX_PATHS.SETTINGS)!;

      // Referenced RSIDs should be present
      for (const rsid of referencedRsids) {
        expect(settingsOut).toContain(rsid);
      }

      // Orphan RSIDs should be gone
      for (const rsid of orphanRsids) {
        expect(settingsOut).not.toContain(rsid);
      }

      // rsidRoot preserved
      expect(settingsOut).toContain(`w:rsidRoot`);
      expect(settingsOut).toContain(rsidRoot);
    });

    it('preserves rsidRoot even if not referenced in document.xml', async () => {
      const rsidRoot = '00FF9999';
      const referencedRsids = ['00AA1111'];
      const allRsids = [rsidRoot, ...referencedRsids, '00BB2222'];

      // Document XML only references 00AA1111, not the rsidRoot
      const docXml = documentWithRsids(referencedRsids);
      const setXml = settingsWithRsids(rsidRoot, allRsids);
      const buffer = await createDocxWithDocAndSettings(docXml, setXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.preserveRawXml();
      doc.stripOrphanRSIDs();
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const settingsOut = zip.getFileAsString(DOCX_PATHS.SETTINGS)!;

      // rsidRoot preserved (spec requirement)
      expect(settingsOut).toContain(rsidRoot);
      // Referenced RSID preserved
      expect(settingsOut).toContain('00AA1111');
      // Orphan removed
      expect(settingsOut).not.toContain('00BB2222');
    });

    it('preserves all RSIDs when all are referenced', async () => {
      const rsidRoot = '00AA1111';
      const allRsids = ['00AA1111', '00BB2222'];

      const docXml = documentWithRsids(allRsids);
      const setXml = settingsWithRsids(rsidRoot, allRsids);
      const buffer = await createDocxWithDocAndSettings(docXml, setXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.preserveRawXml();
      doc.stripOrphanRSIDs();
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const settingsOut = zip.getFileAsString(DOCX_PATHS.SETTINGS)!;

      for (const rsid of allRsids) {
        expect(settingsOut).toContain(rsid);
      }
    });

    it('settings.xml shrinks after stripping many orphans', async () => {
      const rsidRoot = '00000001';
      // Generate 50 RSIDs, only 2 referenced
      const allRsids: string[] = [];
      for (let i = 1; i <= 50; i++) {
        allRsids.push(i.toString(16).toUpperCase().padStart(8, '0'));
      }
      const referencedRsids = [allRsids[0]!, allRsids[1]!];

      const docXml = documentWithRsids(referencedRsids);
      const setXml = settingsWithRsids(rsidRoot, allRsids);
      const buffer = await createDocxWithDocAndSettings(docXml, setXml);

      // Measure original settings size
      const zipOrig = new ZipHandler();
      await zipOrig.loadFromBuffer(buffer);
      const origSettings = zipOrig.getFileAsString(DOCX_PATHS.SETTINGS)!;

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.stripOrphanRSIDs();
      const output = await doc.toBuffer();
      doc.dispose();

      const zipOut = new ZipHandler();
      await zipOut.loadFromBuffer(output);
      const newSettings = zipOut.getFileAsString(DOCX_PATHS.SETTINGS)!;

      // The new settings should be significantly smaller
      expect(newSettings.length).toBeLessThan(origSettings.length);

      // Count w:rsid elements (not w:rsidRoot)
      const origCount = (origSettings.match(/<w:rsid\s/g) || []).length;
      const newCount = (newSettings.match(/<w:rsid\s/g) || []).length;
      expect(newCount).toBeLessThanOrEqual(2); // rsidRoot + referenced
      expect(origCount).toBe(50);
    });
  });

  // =========================================================================
  // Combined sanitization
  // =========================================================================
  describe('Combined sanitization', () => {

    it('both methods chained: doc.flattenFieldCodes().stripOrphanRSIDs()', async () => {
      const referencedRsids = ['00AA1111'];
      const orphanRsids = ['00BB2222', '00CC3333'];
      const allRsids = [...referencedRsids, ...orphanRsids];
      const rsidRoot = '00AA1111';

      // Document with INCLUDEPICTURE + RSID references
      const bodyContent = `
    <w:p w:rsidR="00AA1111" w:rsidRDefault="00AA1111">
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> INCLUDEPICTURE "cid:image001.png" </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>`;

      const docXml = wrapDocXml(bodyContent);
      const setXml = settingsWithRsids(rsidRoot, allRsids);
      const buffer = await createDocxWithDocAndSettings(docXml, setXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes().stripOrphanRSIDs();
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;
      const settingsOut = zip.getFileAsString(DOCX_PATHS.SETTINGS)!;

      // INCLUDEPICTURE stripped, drawing kept
      expect(resultXml).not.toContain('INCLUDEPICTURE');
      expect(resultXml).toContain('w:drawing');

      // Orphan RSIDs removed
      expect(settingsOut).not.toContain('00BB2222');
      expect(settingsOut).not.toContain('00CC3333');
      // Referenced RSID kept
      expect(settingsOut).toContain('00AA1111');
    });

    it('save pipeline order: field flatten runs before RSID scan', async () => {
      // If RSID scan ran before field flatten, it might find RSID references
      // inside the field markup that should be removed. Verify correct order.
      const rsidRoot = '00AA1111';
      const allRsids = ['00AA1111', '00BB2222'];

      // The INCLUDEPICTURE field run has rsidR="00BB2222" — after flattening,
      // 00BB2222 should become orphan if nothing else references it
      const bodyContent = `
    <w:p w:rsidR="00AA1111" w:rsidRDefault="00AA1111">
      <w:r w:rsidR="00BB2222"><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r w:rsidR="00BB2222"><w:instrText xml:space="preserve"> INCLUDEPICTURE "cid:image001.png" </w:instrText></w:r>
      <w:r w:rsidR="00BB2222"><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:inline><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>
      <w:r w:rsidR="00BB2222"><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>`;

      const docXml = wrapDocXml(bodyContent);
      const setXml = settingsWithRsids(rsidRoot, allRsids);
      const buffer = await createDocxWithDocAndSettings(docXml, setXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes().stripOrphanRSIDs();
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const settingsOut = zip.getFileAsString(DOCX_PATHS.SETTINGS)!;

      // 00BB2222 was only referenced on the now-removed field runs → should be orphan
      expect(settingsOut).not.toContain('00BB2222');
      // rsidRoot preserved
      expect(settingsOut).toContain('00AA1111');
    });

    it('round-trip: load, sanitize, save, reload — document renders correctly', async () => {
      const docXml = wrapDocXml(SINGLE_INCLUDEPICTURE + PLAIN_PARAGRAPH);
      const buffer = await createDocxWithDocumentXml(docXml);

      // First pass: sanitize
      const doc1 = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc1.flattenFieldCodes().stripOrphanRSIDs();
      const output1 = await doc1.toBuffer();
      doc1.dispose();

      // Second pass: reload the sanitized document
      const doc2 = await Document.loadFromBuffer(output1, { revisionHandling: 'preserve' });

      // Should load without errors and have content
      const paragraphs = doc2.getParagraphs();
      expect(paragraphs.length).toBeGreaterThan(0);

      // Should be able to save again without issues
      const output2 = await doc2.toBuffer();
      expect(output2.length).toBeGreaterThan(0);

      // Verify the output is a valid ZIP
      const zip = new ZipHandler();
      await zip.loadFromBuffer(output2);
      expect(zip.hasFile(DOCX_PATHS.DOCUMENT)).toBe(true);

      doc2.dispose();
    });
  });
});
