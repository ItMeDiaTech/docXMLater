/**
 * Tests for section properties parsing and round-trip fidelity.
 * Regression tests for inline sectPr serialization and body-level sectPr detection bugs.
 */

import * as fs from 'fs/promises';
import * as path from 'path';
import { Document, Header, Footer } from '../../src';
import { Paragraph } from '../../src/elements/Paragraph';
import { XMLParser } from '../../src/xml/XMLParser';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

const TEST_OUTPUT_DIR = path.join(__dirname, '../../test-output');

describe('Section Properties Parsing', () => {
  beforeAll(async () => {
    await fs.mkdir(TEST_OUTPUT_DIR, { recursive: true });
  });

  afterAll(async () => {
    try {
      await fs.rm(TEST_OUTPUT_DIR, { recursive: true, force: true });
    } catch {
      // Ignore cleanup errors
    }
  });

  describe('Inline sectPr serialization', () => {
    it('should serialize inline sectPr as raw XML passthrough when stored as string', () => {
      const paragraph = new Paragraph();
      const rawSectPr = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';
      paragraph.setSectionProperties(rawSectPr);

      const xml = paragraph.toXML();
      const xmlElements = Array.isArray(xml) ? xml : [xml];

      // Convert to string for inspection
      let output = '';
      for (const el of xmlElements) {
        output += XMLBuilder.elementToString(el);
      }

      // Should contain the raw sectPr XML inside pPr
      expect(output).toContain('<w:sectPr>');
      expect(output).toContain('w:w="12240"');
      expect(output).toContain('w:h="15840"');
      // Should NOT contain malformed attributes like [object Object]
      expect(output).not.toContain('[object Object]');
    });

    it('should not produce malformed XML when sectPr is a parsed object', () => {
      const paragraph = new Paragraph();
      // Simulate what the old parser would store - a parsed object
      paragraph.setSectionProperties({ 'w:pgSz': { '@_w:w': '12240' } });

      const xml = paragraph.toXML();
      const xmlElements = Array.isArray(xml) ? xml : [xml];

      let output = '';
      for (const el of xmlElements) {
        output += XMLBuilder.elementToString(el);
      }

      // The non-string sectPr should be silently skipped to prevent corruption
      expect(output).not.toContain('[object Object]');
    });
  });

  describe('Body-level sectPr detection', () => {
    it('should find body-level sectPr after the last paragraph, not inline ones', () => {
      // Document XML with both inline sectPr (inside pPr) and body-level sectPr
      const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
      </w:pPr>
      <w:r><w:t>Section break here</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>After section break</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId10"/>
      <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
      <w:pgMar w:top="1800" w:right="1440" w:bottom="1800" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

      // Extract body content
      const bodyElements = XMLParser.extractElements(docXml, "w:body");
      expect(bodyElements.length).toBeGreaterThan(0);
      const bodyContent = bodyElements[0]!;

      // Find body-level sectPr using the fixed algorithm
      const lastPClose = bodyContent.lastIndexOf('</w:p>');
      const lastTblClose = bodyContent.lastIndexOf('</w:tbl>');
      const lastSdtClose = bodyContent.lastIndexOf('</w:sdt>');
      const lastBlockEnd = Math.max(lastPClose, lastTblClose, lastSdtClose);

      let bodySectPr: string | undefined;
      if (lastBlockEnd !== -1) {
        const tailContent = bodyContent.substring(lastBlockEnd);
        const sectPrElements = XMLParser.extractElements(tailContent, "w:sectPr");
        if (sectPrElements.length > 0) {
          bodySectPr = sectPrElements[0];
        }
      }

      expect(bodySectPr).toBeDefined();
      // Body-level sectPr should have the landscape page size
      expect(bodySectPr).toContain('w:orient="landscape"');
      // Body-level sectPr should have the header reference
      expect(bodySectPr).toContain('w:headerReference');
      expect(bodySectPr).toContain('rId10');
    });

    it('should handle documents with only body-level sectPr (no inline)', () => {
      const docXml = `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
    </w:sectPr>
  </w:body>
</w:document>`;

      const bodyElements = XMLParser.extractElements(docXml, "w:body");
      const bodyContent = bodyElements[0]!;

      const lastPClose = bodyContent.lastIndexOf('</w:p>');
      const tailContent = bodyContent.substring(lastPClose);
      const sectPrElements = XMLParser.extractElements(tailContent, "w:sectPr");

      expect(sectPrElements.length).toBe(1);
      expect(sectPrElements[0]).toContain('w:w="12240"');
    });
  });

  describe('Header/footer preservation through round-trip', () => {
    it('should preserve header/footer references in section properties', async () => {
      const doc = Document.create();

      // Create and set header and footer
      const header = Header.createDefault();
      header.createParagraph('Test Header');
      const footer = Footer.createDefault();
      footer.createParagraph('Test Footer');
      doc.setHeader(header);
      doc.setFooter(footer);
      doc.createParagraph('Body content');

      // Save and reload
      const buffer = await doc.toBuffer();
      doc.dispose();

      const loaded = await Document.loadFromBuffer(buffer);
      const props = loaded.getSection().getProperties();

      // Headers and footers should be preserved
      expect(props.headers).toBeDefined();
      expect(props.headers?.default).toBeDefined();
      expect(props.footers).toBeDefined();
      expect(props.footers?.default).toBeDefined();

      loaded.dispose();
    });
  });

  describe('Paragraph balance in generated documents', () => {
    it('should produce balanced paragraph tags in output', async () => {
      const doc = Document.create();
      for (let i = 0; i < 10; i++) {
        doc.createParagraph(`Paragraph ${i}`);
      }

      const buffer = await doc.toBuffer();
      doc.dispose();

      // Extract document.xml and check paragraph balance
      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(buffer);
      const docXml = await zip.files['word/document.xml'].async('string');

      const openTags = (docXml.match(/<w:p[\s>]/g) || []).length;
      const closeTags = (docXml.match(/<\/w:p>/g) || []).length;

      expect(openTags).toBe(closeTags);
      expect(openTags).toBeGreaterThanOrEqual(10);
    });
  });
});
