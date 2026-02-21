/**
 * Document Sanitization Tests
 *
 * Tests cover:
 * 1. flattenFieldCodes() — INCLUDEPICTURE field flattening
 * 2. stripOrphanRSIDs() — orphan RSID removal from settings.xml
 * 3. clearDirectSpacingForStyles() — direct w:spacing removal from styled paragraphs
 * 4. Combined sanitization — chained methods
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
            xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
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

      // HYPERLINK field is converted to w:hyperlink element during parsing,
      // so after model regeneration it appears as <w:hyperlink> not raw field code
      expect(resultXml).toContain('w:hyperlink');
      expect(resultXml).toContain('Click here');

      // TOC field should still have its complex field structure
      expect(resultXml).toContain('TOC');
      expect(resultXml).toContain('Table of Contents placeholder');

      // TOC field should have fldChar elements (begin, separate, end)
      const fldCharBeginCount = (resultXml.match(/w:fldCharType\s*=\s*"begin"/g) || []).length;
      const fldCharEndCount = (resultXml.match(/w:fldCharType\s*=\s*"end"/g) || []).length;
      expect(fldCharBeginCount).toBeGreaterThanOrEqual(1);
      expect(fldCharEndCount).toBeGreaterThanOrEqual(1);
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
      // HYPERLINK field preserved (rendered as w:hyperlink element after model regeneration)
      expect(resultXml).toContain('w:hyperlink');
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

      // Other fields preserved (HYPERLINK rendered as w:hyperlink after model regeneration)
      expect(resultXml).toContain('w:hyperlink');
      expect(resultXml).toContain('TOC');
    });

    it('preserves in-memory model changes when flattenFieldCodes is called', async () => {
      // This test verifies the critical fix: flattenFieldCodes() no longer sets
      // skipDocumentXmlRegeneration, so in-memory changes (style application,
      // new paragraphs, etc.) survive the save pipeline alongside field flattening.
      const docXml = wrapDocXml(SINGLE_INCLUDEPICTURE + PLAIN_PARAGRAPH);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Make an in-memory model change: add a new paragraph
      doc.addParagraph(new Paragraph().addText('Added via model'));

      // Call flattenFieldCodes — previously this would set skipDocumentXmlRegeneration = true,
      // causing the in-memory model (including the new paragraph) to be discarded on save
      doc.flattenFieldCodes();
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // The new paragraph added via the model MUST be present in output
      expect(resultXml).toContain('Added via model');

      // The original plain paragraph content should also survive
      expect(resultXml).toContain('Hello World');

      // INCLUDEPICTURE field markup should be stripped by post-processing
      expect(resultXml).not.toContain('INCLUDEPICTURE');

      // Drawing content should be preserved (via parser round-trip or post-processing)
      expect(resultXml).toContain('w:drawing');
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
  // clearDirectSpacingForStyles()
  // =========================================================================
  describe('clearDirectSpacingForStyles()', () => {

    it('removes direct w:spacing from paragraphs with matching style', async () => {
      const bodyContent = `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
        <w:spacing w:before="120" w:after="120"/>
      </w:pPr>
      <w:r><w:t>Hello</w:t></w:r>
    </w:p>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes();
      doc.clearDirectSpacingForStyles(['Normal']);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // w:spacing should be removed from the Normal paragraph
      expect(resultXml).toContain('<w:pStyle w:val="Normal"/>');
      expect(resultXml).not.toMatch(/<w:spacing\s+w:before="120"/);
      // Text preserved
      expect(resultXml).toContain('Hello');
    });

    it('leaves paragraphs with non-matching styles unchanged', async () => {
      const bodyContent = `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="CustomStyle"/>
        <w:spacing w:before="240" w:after="240"/>
      </w:pPr>
      <w:r><w:t>Custom</w:t></w:r>
    </w:p>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes();
      doc.clearDirectSpacingForStyles(['Normal', 'Heading1']);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // CustomStyle is not in the target list — spacing should remain
      expect(resultXml).toContain('w:before="240"');
      expect(resultXml).toContain('w:after="240"');
    });

    it('leaves paragraphs without w:pStyle unchanged', async () => {
      const bodyContent = `
    <w:p>
      <w:pPr>
        <w:spacing w:before="100" w:after="100"/>
      </w:pPr>
      <w:r><w:t>No style</w:t></w:r>
    </w:p>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes();
      doc.clearDirectSpacingForStyles(['Normal']);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // No pStyle — spacing should remain (paragraph needs direct formatting)
      expect(resultXml).toContain('w:before="100"');
    });

    it('preserves w:spacing inside w:pPrChange (revision history)', async () => {
      const bodyContent = `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
        <w:spacing w:before="120" w:after="120"/>
        <w:pPrChange w:id="1" w:author="John" w:date="2024-01-01T00:00:00Z">
          <w:pPr>
            <w:pStyle w:val="Normal"/>
            <w:spacing w:before="240" w:after="240"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>Tracked</w:t></w:r>
    </w:p>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes();
      doc.clearDirectSpacingForStyles(['Normal']);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Direct spacing removed
      expect(resultXml).not.toMatch(/<w:pPr>\s*<w:pStyle w:val="Normal"\/>[\s\S]*?<w:spacing w:before="120"/);
      // Historical spacing inside pPrChange preserved
      expect(resultXml).toContain('w:before="240"');
      expect(resultXml).toContain('w:after="240"');
      // pPrChange structure intact
      expect(resultXml).toContain('<w:pPrChange');
      expect(resultXml).toContain('</w:pPrChange>');
    });

    it('preserves w:spacing inside w:rPr (paragraph mark character spacing)', async () => {
      const bodyContent = `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
        <w:spacing w:before="120" w:after="120"/>
        <w:rPr>
          <w:spacing w:val="20"/>
          <w:b/>
        </w:rPr>
      </w:pPr>
      <w:r><w:t>With rPr</w:t></w:r>
    </w:p>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.clearDirectSpacingForStyles(['Normal']);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Direct paragraph spacing removed by clearDirectSpacingForStyles
      expect(resultXml).not.toContain('w:before="120"');
      // Character spacing inside rPr preserved
      expect(resultXml).toContain('w:val="20"');
      // rPr structure intact (model regenerates bold as w:val="1")
      expect(resultXml).toContain('w:rPr');
      expect(resultXml).toMatch(/w:b/);
    });

    it('handles table cell paragraphs with matching styles', async () => {
      const bodyContent = `
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:pPr>
              <w:pStyle w:val="Normal"/>
              <w:spacing w:before="120" w:after="120"/>
            </w:pPr>
            <w:r><w:t>Cell text</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes();
      doc.clearDirectSpacingForStyles(['Normal']);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Table cell paragraph spacing removed
      expect(resultXml).toContain('<w:pStyle w:val="Normal"/>');
      expect(resultXml).not.toContain('w:before="120"');
      expect(resultXml).toContain('Cell text');
    });

    it('handles multiple paragraphs with mixed styles', async () => {
      const bodyContent = `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
        <w:spacing w:before="120" w:after="120"/>
      </w:pPr>
      <w:r><w:t>Normal para</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:spacing w:before="240" w:after="60"/>
      </w:pPr>
      <w:r><w:t>Heading</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="CustomStyle"/>
        <w:spacing w:before="300" w:after="300"/>
      </w:pPr>
      <w:r><w:t>Custom</w:t></w:r>
    </w:p>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes();
      doc.clearDirectSpacingForStyles(['Normal', 'Heading1']);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Normal and Heading1 spacing removed
      expect(resultXml).not.toContain('w:before="120"');
      expect(resultXml).not.toContain('w:before="240"');
      // CustomStyle spacing preserved
      expect(resultXml).toContain('w:before="300"');
      expect(resultXml).toContain('w:after="300"');
    });

    it('is a no-op when called with empty style list', async () => {
      const bodyContent = `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
        <w:spacing w:before="120" w:after="120"/>
      </w:pPr>
      <w:r><w:t>Hello</w:t></w:r>
    </w:p>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.flattenFieldCodes();
      doc.clearDirectSpacingForStyles([]);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Nothing should change
      expect(resultXml).toContain('w:before="120"');
      expect(resultXml).toContain('w:after="120"');
    });

    it('does not corrupt XML when pPrChange child pPr has matching style', async () => {
      // Regression: outer pPr has pStyle="ListParagraph" + spacing + pPrChange
      // containing inner pPr with pStyle="ListParagraph" + spacing.
      // Bug: searchPos only advanced past the opening tag, causing the inner pPr
      // to be found as a separate top-level match, producing orphaned XML fragments.
      const bodyContent = `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="ListParagraph"/>
        <w:spacing w:before="120" w:after="120"/>
        <w:pPrChange w:id="10" w:author="Alice" w:date="2025-01-01T00:00:00Z">
          <w:pPr>
            <w:pStyle w:val="ListParagraph"/>
            <w:spacing w:before="240" w:after="240"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>List item</w:t></w:r>
    </w:p>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.preserveRawXml();
      doc.clearDirectSpacingForStyles(['ListParagraph']);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Verify output is well-formed: no orphaned tags
      const openCount = (resultXml.match(/<w:pPr>/g) || []).length;
      const closeCount = (resultXml.match(/<\/w:pPr>/g) || []).length;
      expect(openCount).toBe(closeCount);

      // Direct spacing (before="120") removed from outer pPr
      expect(resultXml).not.toContain('w:before="120"');
      // Historical spacing inside pPrChange preserved
      expect(resultXml).toContain('w:before="240"');
      expect(resultXml).toContain('w:after="240"');
      // pPrChange structure intact
      expect(resultXml).toContain('<w:pPrChange');
      expect(resultXml).toContain('</w:pPrChange>');
      // Text preserved
      expect(resultXml).toContain('List item');
    });

    it('does not match style from pPrChange when outer pPr has different style', async () => {
      // Bug 2: style check searched raw inner content including pPrChange.
      // Outer pPr has pStyle="Normal" (not targeted), but pPrChange has
      // pStyle="ListParagraph" (targeted). Should NOT remove spacing.
      const bodyContent = `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
        <w:spacing w:before="120" w:after="120"/>
        <w:pPrChange w:id="11" w:author="Bob" w:date="2025-02-01T00:00:00Z">
          <w:pPr>
            <w:pStyle w:val="ListParagraph"/>
            <w:spacing w:before="240" w:after="240"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>Was ListParagraph, now Normal</w:t></w:r>
    </w:p>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.preserveRawXml();
      doc.clearDirectSpacingForStyles(['ListParagraph']);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Spacing should NOT be removed — current style is Normal, not ListParagraph
      expect(resultXml).toContain('w:before="120"');
      expect(resultXml).toContain('w:after="120"');
      // pPrChange preserved
      expect(resultXml).toContain('<w:pPrChange');
    });

    it('does not match style from pPrChange when outer pPr has no direct style', async () => {
      // Outer pPr has only rPr + pPrChange with pStyle="ListParagraph".
      // Target: ["ListParagraph"]. No style in outer region → no modification.
      const bodyContent = `
    <w:p>
      <w:pPr>
        <w:spacing w:before="100" w:after="100"/>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:pPrChange w:id="12" w:author="Carol" w:date="2025-03-01T00:00:00Z">
          <w:pPr>
            <w:pStyle w:val="ListParagraph"/>
            <w:spacing w:before="200" w:after="200"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>No current style</w:t></w:r>
    </w:p>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.preserveRawXml();
      doc.clearDirectSpacingForStyles(['ListParagraph']);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // No modification — spacing should remain
      expect(resultXml).toContain('w:before="100"');
      expect(resultXml).toContain('w:after="100"');
      // XML well-formed
      const openCount = (resultXml.match(/<w:pPr>/g) || []).length;
      const closeCount = (resultXml.match(/<\/w:pPr>/g) || []).length;
      expect(openCount).toBe(closeCount);
    });

    it('handles multiple consecutive paragraphs with pPrChange safely', async () => {
      // Two paragraphs both having outer pPr with pStyle="ListParagraph" +
      // spacing + pPrChange. Both should have spacing removed without corruption.
      const bodyContent = `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="ListParagraph"/>
        <w:spacing w:before="120" w:after="120"/>
        <w:pPrChange w:id="20" w:author="Dave" w:date="2025-04-01T00:00:00Z">
          <w:pPr>
            <w:pStyle w:val="ListParagraph"/>
            <w:spacing w:before="240" w:after="240"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>First item</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="ListParagraph"/>
        <w:spacing w:before="150" w:after="150"/>
        <w:pPrChange w:id="21" w:author="Dave" w:date="2025-04-01T00:00:00Z">
          <w:pPr>
            <w:pStyle w:val="ListParagraph"/>
            <w:spacing w:before="300" w:after="300"/>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>Second item</w:t></w:r>
    </w:p>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.preserveRawXml();
      doc.clearDirectSpacingForStyles(['ListParagraph']);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Direct spacing removed from both paragraphs
      expect(resultXml).not.toContain('w:before="120"');
      expect(resultXml).not.toContain('w:before="150"');
      // Historical spacing inside pPrChange preserved
      expect(resultXml).toContain('w:before="240"');
      expect(resultXml).toContain('w:before="300"');
      // XML well-formed: pPr open/close counts match
      const openCount = (resultXml.match(/<w:pPr>/g) || []).length;
      const closeCount = (resultXml.match(/<\/w:pPr>/g) || []).length;
      expect(openCount).toBe(closeCount);
      // Text preserved
      expect(resultXml).toContain('First item');
      expect(resultXml).toContain('Second item');
    });

    it('works without flattenFieldCodes (standalone usage)', async () => {
      const bodyContent = `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
        <w:spacing w:before="120" w:after="120"/>
      </w:pPr>
      <w:r><w:t>Standalone</w:t></w:r>
    </w:p>`;
      const docXml = wrapDocXml(bodyContent);
      const buffer = await createDocxWithDocumentXml(docXml);

      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      doc.preserveRawXml();
      doc.clearDirectSpacingForStyles(['Normal']);
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Spacing should be removed even without flattenFieldCodes
      expect(resultXml).not.toContain('w:before="120"');
      expect(resultXml).toContain('Standalone');
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

    it('round-trip WITHOUT flattenFieldCodes preserves full INCLUDEPICTURE field structure', async () => {
      const docXml = wrapDocXml(SINGLE_INCLUDEPICTURE);
      const buffer = await createDocxWithDocumentXml(docXml);

      // Load and save without calling flattenFieldCodes()
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      const output = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(output);
      const resultXml = zip.getFileAsString(DOCX_PATHS.DOCUMENT)!;

      // Both drawing content AND INCLUDEPICTURE field markup should survive
      expect(resultXml).toContain('w:drawing');
      expect(resultXml).toContain('INCLUDEPICTURE');
      expect(resultXml).toContain('w:fldChar');
      expect(resultXml).toContain('w:instrText');
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
