/**
 * Tests for webSettings.xml round-trip support and sanitization API
 * Covers: preservation, getWebSettingsInfo, stripWebDivs, sanitizeWebSettings, setters, dispose
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { DOCX_PATHS } from '../../src/zip/types';

describe('webSettings.xml Round-Trip', () => {
  /**
   * Helper: create a minimal DOCX buffer with custom webSettings.xml
   */
  async function createDocxWithWebSettings(webSettingsXml: string): Promise<Buffer> {
    const doc = Document.create();
    doc.createParagraph('Test content');
    const buffer = await doc.toBuffer();
    doc.dispose();

    // Post-process: inject custom webSettings.xml
    const zip = new ZipHandler();
    await zip.loadFromBuffer(buffer);
    zip.updateFile(DOCX_PATHS.WEB_SETTINGS, webSettingsXml);
    const modifiedBuffer = await zip.toBuffer();
    zip.clear();

    return modifiedBuffer;
  }

  const minimalWebSettings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:optimizeForBrowser/>
  <w:allowPNG/>
</w:webSettings>`;

  const webSettingsWithDivs = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:divs>
    <w:div w:id="1">
      <w:bodyDiv w:val="1"/>
      <w:marLeft w:val="0"/>
      <w:marRight w:val="0"/>
      <w:marTop w:val="0"/>
      <w:marBottom w:val="0"/>
    </w:div>
    <w:div w:id="2">
      <w:bodyDiv w:val="1"/>
      <w:marLeft w:val="0"/>
      <w:marRight w:val="0"/>
      <w:marTop w:val="0"/>
      <w:marBottom w:val="0"/>
    </w:div>
    <w:div w:id="3">
      <w:bodyDiv w:val="1"/>
    </w:div>
  </w:divs>
  <w:optimizeForBrowser/>
  <w:allowPNG/>
</w:webSettings>`;

  const webSettingsWithFlags = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:optimizeForBrowser/>
  <w:allowPNG/>
  <w:relyOnVML/>
  <w:doNotRelyOnCSS/>
  <w:pixelsPerInch w:val="96"/>
  <w:targetScreenSz w:val="1024x768"/>
  <w:encoding w:val="windows-1252"/>
</w:webSettings>`;

  // =========================================================================
  // New document behavior
  // =========================================================================

  describe('New document behavior', () => {
    it('should include webSettings.xml with optimizeForBrowser and allowPNG', async () => {
      const doc = Document.create();
      doc.createParagraph('Hello');
      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const xml = zip.getFileAsString(DOCX_PATHS.WEB_SETTINGS);
      zip.clear();

      expect(xml).toBeDefined();
      expect(xml).toContain('<w:optimizeForBrowser/>');
      expect(xml).toContain('<w:allowPNG/>');
    });

    it('should have no w:divs in new documents', async () => {
      const doc = Document.create();
      doc.createParagraph('Hello');
      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const xml = zip.getFileAsString(DOCX_PATHS.WEB_SETTINGS);
      zip.clear();

      expect(xml).not.toContain('<w:divs');
      expect(xml).not.toContain('<w:div ');
    });

    it('should report correct defaults from getWebSettingsInfo()', () => {
      const doc = Document.create();
      const info = doc.getWebSettingsInfo();
      doc.dispose();

      expect(info.divCount).toBe(0);
      expect(info.optimizeForBrowser).toBe(true);
      expect(info.allowPNG).toBe(true);
      expect(info.relyOnVML).toBe(false);
      expect(info.doNotRelyOnCSS).toBe(false);
      expect(info.doNotSaveAsSingleFile).toBe(false);
      expect(info.doNotOrganizeInFolder).toBe(false);
      expect(info.doNotUseLongFileNames).toBe(false);
      expect(info.pixelsPerInch).toBeUndefined();
      expect(info.targetScreenSz).toBeUndefined();
      expect(info.encoding).toBeUndefined();
    });
  });

  // =========================================================================
  // Load preservation
  // =========================================================================

  describe('Load preservation', () => {
    it('should preserve webSettings.xml exactly when unmodified', async () => {
      const buffer = await createDocxWithWebSettings(minimalWebSettings);
      const doc = await Document.loadFromBuffer(buffer);

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const savedXml = zip.getFileAsString(DOCX_PATHS.WEB_SETTINGS);
      zip.clear();

      expect(savedXml).toBe(minimalWebSettings);
    });

    it('should preserve w:divs through round-trip when unmodified', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithDivs);
      const doc = await Document.loadFromBuffer(buffer);

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const savedXml = zip.getFileAsString(DOCX_PATHS.WEB_SETTINGS);
      zip.clear();

      expect(savedXml).toBe(webSettingsWithDivs);
      expect(savedXml).toContain('<w:divs>');
    });

    it('should not accumulate divs through two-level round-trip', async () => {
      const buffer1 = await createDocxWithWebSettings(webSettingsWithDivs);
      const doc1 = await Document.loadFromBuffer(buffer1);
      const buffer2 = await doc1.toBuffer();
      doc1.dispose();

      const doc2 = await Document.loadFromBuffer(buffer2);
      const info = doc2.getWebSettingsInfo();
      const buffer3 = await doc2.toBuffer();
      doc2.dispose();

      expect(info.divCount).toBe(3);

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer3);
      const savedXml = zip.getFileAsString(DOCX_PATHS.WEB_SETTINGS);
      zip.clear();

      const matches = savedXml!.match(/<w:div\b/g);
      expect(matches?.length).toBe(3);
    });
  });

  // =========================================================================
  // getWebSettingsInfo()
  // =========================================================================

  describe('getWebSettingsInfo()', () => {
    it('should return divCount=0 for new documents', () => {
      const doc = Document.create();
      expect(doc.getWebSettingsInfo().divCount).toBe(0);
      doc.dispose();
    });

    it('should return correct divCount for documents with w:divs', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithDivs);
      const doc = await Document.loadFromBuffer(buffer);

      expect(doc.getWebSettingsInfo().divCount).toBe(3);
      doc.dispose();
    });

    it('should parse boolean flags correctly', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithFlags);
      const doc = await Document.loadFromBuffer(buffer);
      const info = doc.getWebSettingsInfo();
      doc.dispose();

      expect(info.optimizeForBrowser).toBe(true);
      expect(info.allowPNG).toBe(true);
      expect(info.relyOnVML).toBe(true);
      expect(info.doNotRelyOnCSS).toBe(true);
      expect(info.doNotSaveAsSingleFile).toBe(false);
      expect(info.doNotOrganizeInFolder).toBe(false);
      expect(info.doNotUseLongFileNames).toBe(false);
    });

    it('should parse pixelsPerInch attribute', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithFlags);
      const doc = await Document.loadFromBuffer(buffer);
      const info = doc.getWebSettingsInfo();
      doc.dispose();

      expect(info.pixelsPerInch).toBe(96);
    });

    it('should parse targetScreenSz and encoding', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithFlags);
      const doc = await Document.loadFromBuffer(buffer);
      const info = doc.getWebSettingsInfo();
      doc.dispose();

      expect(info.targetScreenSz).toBe('1024x768');
      expect(info.encoding).toBe('windows-1252');
    });
  });

  // =========================================================================
  // stripWebDivs()
  // =========================================================================

  describe('stripWebDivs()', () => {
    it('should remove w:divs section and return count', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithDivs);
      const doc = await Document.loadFromBuffer(buffer);

      const removed = doc.stripWebDivs();
      expect(removed).toBe(3);

      const info = doc.getWebSettingsInfo();
      expect(info.divCount).toBe(0);
      doc.dispose();
    });

    it('should return 0 when no divs present', async () => {
      const buffer = await createDocxWithWebSettings(minimalWebSettings);
      const doc = await Document.loadFromBuffer(buffer);

      expect(doc.stripWebDivs()).toBe(0);
      doc.dispose();
    });

    it('should return 0 for new documents', () => {
      const doc = Document.create();
      expect(doc.stripWebDivs()).toBe(0);
      doc.dispose();
    });

    it('should not affect other webSettings flags', async () => {
      const xmlWithDivsAndFlags = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:divs>
    <w:div w:id="1"><w:bodyDiv w:val="1"/></w:div>
  </w:divs>
  <w:optimizeForBrowser/>
  <w:allowPNG/>
  <w:relyOnVML/>
</w:webSettings>`;

      const buffer = await createDocxWithWebSettings(xmlWithDivsAndFlags);
      const doc = await Document.loadFromBuffer(buffer);

      doc.stripWebDivs();
      const info = doc.getWebSettingsInfo();
      expect(info.optimizeForBrowser).toBe(true);
      expect(info.allowPNG).toBe(true);
      expect(info.relyOnVML).toBe(true);
      doc.dispose();
    });

    it('should produce clean XML after stripping and saving', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithDivs);
      const doc = await Document.loadFromBuffer(buffer);

      doc.stripWebDivs();
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const savedXml = zip.getFileAsString(DOCX_PATHS.WEB_SETTINGS);
      zip.clear();

      expect(savedXml).not.toContain('<w:divs');
      expect(savedXml).not.toContain('<w:div ');
      expect(savedXml).toContain('<w:optimizeForBrowser/>');
      expect(savedXml).toContain('<w:allowPNG/>');
    });
  });

  // =========================================================================
  // sanitizeWebSettings()
  // =========================================================================

  describe('sanitizeWebSettings()', () => {
    it('should return div count and produce minimal webSettings', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithDivs);
      const doc = await Document.loadFromBuffer(buffer);

      const removed = doc.sanitizeWebSettings();
      expect(removed).toBe(3);

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(outputBuffer);
      const savedXml = zip.getFileAsString(DOCX_PATHS.WEB_SETTINGS);
      zip.clear();

      expect(savedXml).toContain('<w:optimizeForBrowser/>');
      expect(savedXml).toContain('<w:allowPNG/>');
      expect(savedXml).not.toContain('<w:divs');
    });

    it('should reset boolean flags to defaults', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithFlags);
      const doc = await Document.loadFromBuffer(buffer);

      // Verify flags were parsed
      expect(doc.getWebSettingsInfo().relyOnVML).toBe(true);

      doc.sanitizeWebSettings();
      const info = doc.getWebSettingsInfo();
      expect(info.relyOnVML).toBe(false);
      expect(info.doNotRelyOnCSS).toBe(false);
      expect(info.optimizeForBrowser).toBe(true);
      expect(info.allowPNG).toBe(true);
      doc.dispose();
    });

    it('should clear attributes (pixelsPerInch, encoding)', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithFlags);
      const doc = await Document.loadFromBuffer(buffer);

      doc.sanitizeWebSettings();
      const info = doc.getWebSettingsInfo();
      expect(info.pixelsPerInch).toBeUndefined();
      expect(info.targetScreenSz).toBeUndefined();
      expect(info.encoding).toBeUndefined();
      doc.dispose();
    });

    it('should produce output matching new document template', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithDivs);
      const doc = await Document.loadFromBuffer(buffer);

      doc.sanitizeWebSettings();
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      // Create a brand new document for comparison
      const newDoc = Document.create();
      newDoc.createParagraph('Comparison');
      const newBuffer = await newDoc.toBuffer();
      newDoc.dispose();

      const zip1 = new ZipHandler();
      await zip1.loadFromBuffer(outputBuffer);
      const sanitizedXml = zip1.getFileAsString(DOCX_PATHS.WEB_SETTINGS);
      zip1.clear();

      const zip2 = new ZipHandler();
      await zip2.loadFromBuffer(newBuffer);
      const newDocXml = zip2.getFileAsString(DOCX_PATHS.WEB_SETTINGS);
      zip2.clear();

      expect(sanitizedXml).toBe(newDocXml);
    });
  });

  // =========================================================================
  // Boolean flag setters
  // =========================================================================

  describe('Boolean flag setters', () => {
    it('should remove optimizeForBrowser when set to false', async () => {
      const doc = Document.create();
      doc.createParagraph('Test');
      doc.setOptimizeForBrowser(false);

      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const xml = zip.getFileAsString(DOCX_PATHS.WEB_SETTINGS);
      zip.clear();

      expect(xml).not.toContain('<w:optimizeForBrowser');
      expect(xml).toContain('<w:allowPNG/>');
    });

    it('should remove allowPNG when set to false', async () => {
      const doc = Document.create();
      doc.createParagraph('Test');
      doc.setAllowPNG(false);

      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const xml = zip.getFileAsString(DOCX_PATHS.WEB_SETTINGS);
      zip.clear();

      expect(xml).not.toContain('<w:allowPNG');
      expect(xml).toContain('<w:optimizeForBrowser/>');
    });

    it('should support chaining (return document instance)', () => {
      const doc = Document.create();
      const result = doc.setOptimizeForBrowser(true).setAllowPNG(false);
      expect(result).toBe(doc);
      expect(doc.getOptimizeForBrowser()).toBe(true);
      expect(doc.getAllowPNG()).toBe(false);
      doc.dispose();
    });
  });

  // =========================================================================
  // dispose()
  // =========================================================================

  describe('dispose()', () => {
    it('should clear webSettings state', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithDivs);
      const doc = await Document.loadFromBuffer(buffer);

      expect(doc.getWebSettingsInfo().divCount).toBe(3);

      doc.dispose();

      // After dispose, state should be reset to defaults
      const info = doc.getWebSettingsInfo();
      expect(info.divCount).toBe(0);
    });

    it('should reset webSettings to defaults after dispose', async () => {
      const buffer = await createDocxWithWebSettings(webSettingsWithFlags);
      const doc = await Document.loadFromBuffer(buffer);

      expect(doc.getWebSettingsInfo().relyOnVML).toBe(true);
      expect(doc.getWebSettingsInfo().pixelsPerInch).toBe(96);

      doc.dispose();

      const info = doc.getWebSettingsInfo();
      expect(info.optimizeForBrowser).toBe(true);
      expect(info.allowPNG).toBe(true);
      expect(info.relyOnVML).toBe(false);
      expect(info.pixelsPerInch).toBeUndefined();
    });
  });
});
