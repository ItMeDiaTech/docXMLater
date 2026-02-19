/**
 * Header/Footer Round-Trip Test
 *
 * Tests that headers and footers are correctly:
 * 1. Parsed when loading a document
 * 2. Saved with correct [Content_Types].xml entries
 * 3. Preserved through load/save cycles
 */

import { Document, Header, Footer } from '../../src';
import * as path from 'path';
import * as fs from 'fs';

describe('Header/Footer Round-Trip', () => {
  const outputDir = path.join(__dirname, '../output');

  beforeAll(() => {
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }
  });

  it('should create, save, and load document with headers/footers', async () => {
    // Step 1: Create a document with header and footer
    const doc1 = Document.create({
      properties: {
        title: 'Header/Footer Round-Trip Test',
        creator: 'DocXMLater Test Suite',
      },
    });

    // Create header with text
    const header = Header.createDefault();
    header.createParagraph('Document Header').setAlignment('right');

    // Create footer with text
    const footer = Footer.createDefault();
    footer.createParagraph('Page Footer').setAlignment('center');

    // Set header and footer
    doc1.setHeader(header);
    doc1.setFooter(footer);

    // Add body content
    doc1.createParagraph('This document has a header and footer.');

    // Save the document
    const testPath = path.join(outputDir, 'test-header-footer-roundtrip.docx');
    await doc1.save(testPath);

    expect(fs.existsSync(testPath)).toBe(true);

    // Step 2: Load the document back
    const doc2 = await Document.load(testPath);

    // Verify headers and footers were loaded
    const headerFooterManager = doc2.getHeaderFooterManager();
    expect(headerFooterManager.getHeaderCount()).toBeGreaterThan(0);
    expect(headerFooterManager.getFooterCount()).toBeGreaterThan(0);

    // Step 3: Save the loaded document again
    const testPath2 = path.join(outputDir, 'test-header-footer-roundtrip2.docx');
    await doc2.save(testPath2);

    expect(fs.existsSync(testPath2)).toBe(true);

    // Step 4: Verify [Content_Types].xml has correct entries
    // Load the saved document as ZIP and check Content_Types
    const JSZip = require('jszip');
    const fileBuffer = fs.readFileSync(testPath2);
    const zip = await JSZip.loadAsync(fileBuffer);

    const contentTypesFile = zip.file('[Content_Types].xml');
    expect(contentTypesFile).not.toBeNull();

    const contentTypes = await contentTypesFile!.async('string');

    // Verify Override entries for header and footer
    expect(contentTypes).toContain('application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
    expect(contentTypes).toContain('application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');

    // Verify the header/footer XML files exist
    expect(zip.file('word/header1.xml')).not.toBeNull();
    expect(zip.file('word/footer1.xml')).not.toBeNull();
  });

  it('should handle remove header/footer operations', async () => {
    // Create document with header and footer
    const doc = Document.create();

    const header = Header.createDefault();
    header.createParagraph('Test Header');

    const footer = Footer.createDefault();
    footer.createParagraph('Test Footer');

    doc.setHeader(header);
    doc.setFooter(footer);

    doc.createParagraph('Content with header and footer');

    // Save initial version
    const testPath1 = path.join(outputDir, 'test-remove-header-footer1.docx');
    await doc.save(testPath1);

    // Remove header
    doc.removeHeader('default');

    // Save after removing header
    const testPath2 = path.join(outputDir, 'test-remove-header-footer2.docx');
    await doc.save(testPath2);

    // Verify header is removed but footer still exists
    const JSZip = require('jszip');
    const fileBuffer2 = fs.readFileSync(testPath2);
    const zip2 = await JSZip.loadAsync(fileBuffer2);

    // Footer should still exist
    expect(zip2.file('word/footer1.xml')).not.toBeNull();

    // Content Types should still have footer override
    const contentTypes2 = await zip2.file('[Content_Types].xml')!.async('string');
    expect(contentTypes2).toContain('application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
  });

  it('should handle clearHeaders and clearFooters operations', async () => {
    // Create document with multiple headers and footers
    const doc = Document.create();

    // Set first page header/footer
    const firstHeader = Header.createFirst();
    firstHeader.createParagraph('First Page Header');

    const firstFooter = Footer.createFirst();
    firstFooter.createParagraph('First Page Footer');

    // Set default header/footer
    const header = Header.createDefault();
    header.createParagraph('Default Header');

    const footer = Footer.createDefault();
    footer.createParagraph('Default Footer');

    doc.setFirstPageHeader(firstHeader);
    doc.setHeader(header);
    doc.setFirstPageFooter(firstFooter);
    doc.setFooter(footer);

    doc.createParagraph('Content');

    // Clear all headers
    doc.clearHeaders();

    // Save after clearing headers
    const testPath = path.join(outputDir, 'test-clear-headers.docx');
    await doc.save(testPath);

    // Verify headers are removed but footers still exist
    const JSZip = require('jszip');
    const fileBuffer = fs.readFileSync(testPath);
    const zip = await JSZip.loadAsync(fileBuffer);

    // Footers should still exist
    expect(zip.file('word/footer1.xml')).not.toBeNull();

    // Clear footers
    doc.clearFooters();

    // Save after clearing footers
    const testPath2 = path.join(outputDir, 'test-clear-all.docx');
    await doc.save(testPath2);

    // Verify no header/footer files exist
    const fileBuffer2 = fs.readFileSync(testPath2);
    const zip2 = await JSZip.loadAsync(fileBuffer2);

    const contentTypes2 = await zip2.file('[Content_Types].xml')!.async('string');

    // Should not contain header/footer content types if no headers/footers exist
    // (or should handle gracefully if empty headers/footers are saved)
  });

  it('should clear header/footer content while preserving structure', async () => {
    // Create document with headers and footers
    const doc = Document.create();
    const header = Header.createDefault();
    header.createParagraph('Important Header Text');
    const footer = Footer.createDefault();
    footer.createParagraph('Page 1 of 10');
    doc.setHeader(header);
    doc.setFooter(footer);
    doc.createParagraph('Body content');

    // Clear content (not remove)
    const count = doc.clearAllHeaderFooterContent();
    expect(count).toBe(2);

    const buffer = await doc.toBuffer();
    doc.dispose();

    // Verify: header/footer files still exist but are empty
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(buffer);

    expect(zip.file('word/header1.xml')).not.toBeNull();
    expect(zip.file('word/footer1.xml')).not.toBeNull();

    const headerXml = await zip.file('word/header1.xml')!.async('string');
    const footerXml = await zip.file('word/footer1.xml')!.async('string');

    // Content should be gone
    expect(headerXml).not.toContain('Important Header Text');
    expect(footerXml).not.toContain('Page 1 of 10');

    // Structure should remain (empty paragraph)
    expect(headerXml).toContain('w:hdr');
    expect(headerXml).toContain('w:p');
    expect(footerXml).toContain('w:ftr');
    expect(footerXml).toContain('w:p');

    // Relationships should still exist
    const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(relsXml).toContain('/header');
    expect(relsXml).toContain('/footer');
  });

  it('should clear loaded header/footer content (rawXML reset)', async () => {
    // Create doc with header content
    const doc1 = Document.create();
    const header = Header.createDefault();
    header.createParagraph('Original Header');
    const footer = Footer.createDefault();
    footer.createParagraph('Original Footer');
    doc1.setHeader(header);
    doc1.setFooter(footer);
    doc1.createParagraph('Body');
    const buffer1 = await doc1.toBuffer();
    doc1.dispose();

    // Load and clear
    const doc2 = await Document.loadFromBuffer(buffer1);
    doc2.clearAllHeaderFooterContent();
    const buffer2 = await doc2.toBuffer();
    doc2.dispose();

    // Verify the rawXML was reset — content should be empty
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(buffer2);

    const headerXml = await zip.file('word/header1.xml')!.async('string');
    expect(headerXml).not.toContain('Original Header');
    expect(headerXml).toContain('w:hdr');

    const footerXml = await zip.file('word/footer1.xml')!.async('string');
    expect(footerXml).not.toContain('Original Footer');
    expect(footerXml).toContain('w:ftr');
  });

  it('should remove all headers/footers including from inline sectPr in multi-section documents', async () => {
    // Create a document with a section break that has headers/footers in the inline sectPr
    const doc = Document.create();

    // Set up default header and footer for the body-level section
    const header = Header.createDefault();
    header.createParagraph('Body Header');
    const footer = Footer.createDefault();
    footer.createParagraph('Body Footer');
    doc.setHeader(header);
    doc.setFooter(footer);

    // Create a paragraph with an inline sectPr containing header/footer references
    // This simulates a multi-section document where the first section has its own headers/footers
    const para = doc.createParagraph('Section 1 content');
    para.setSectionProperties(
      '<w:sectPr>' +
      '<w:headerReference w:type="default" r:id="rId99"/>' +
      '<w:footerReference w:type="default" r:id="rId100"/>' +
      '<w:titlePg/>' +
      '<w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
      '</w:sectPr>'
    );

    doc.createParagraph('Section 2 content');

    // Remove all headers/footers
    doc.removeAllHeadersFooters();

    // Save the document
    const buffer = await doc.toBuffer();
    doc.dispose();

    // Verify: no header/footer files exist in the ZIP
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(buffer);

    const headerFiles = Object.keys(zip.files).filter(f => f.match(/^word\/header\d+\.xml$/));
    const footerFiles = Object.keys(zip.files).filter(f => f.match(/^word\/footer\d+\.xml$/));
    expect(headerFiles).toHaveLength(0);
    expect(footerFiles).toHaveLength(0);

    // Verify: no headerReference or footerReference in the document XML
    const docXml = await zip.file('word/document.xml')!.async('string');
    expect(docXml).not.toContain('w:headerReference');
    expect(docXml).not.toContain('w:footerReference');
    expect(docXml).not.toContain('w:titlePg');

    // Verify: the inline sectPr still has page size/margins (not completely removed)
    expect(docXml).toContain('w:pgSz');
    expect(docXml).toContain('w:pgMar');
  });

  it('should remove headers/footers from loaded multi-section document round-trip', async () => {
    // Build a minimal multi-section document with headers/footers
    const doc1 = Document.create();

    const header = Header.createDefault();
    header.createParagraph('Default Header');
    const footer = Footer.createDefault();
    footer.createParagraph('Default Footer');
    doc1.setHeader(header);
    doc1.setFooter(footer);

    doc1.createParagraph('Page 1');
    const buffer1 = await doc1.toBuffer();
    doc1.dispose();

    // Load the document, remove headers/footers, save again
    const doc2 = await Document.loadFromBuffer(buffer1);
    const removedCount = doc2.removeAllHeadersFooters();
    expect(removedCount).toBeGreaterThan(0);

    const buffer2 = await doc2.toBuffer();
    doc2.dispose();

    // Verify the result has no headers/footers
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(buffer2);

    const headerFiles = Object.keys(zip.files).filter(f => f.match(/^word\/header\d+\.xml$/));
    const footerFiles = Object.keys(zip.files).filter(f => f.match(/^word\/footer\d+\.xml$/));
    expect(headerFiles).toHaveLength(0);
    expect(footerFiles).toHaveLength(0);

    // Verify no header/footer references in the document XML
    const docXml = await zip.file('word/document.xml')!.async('string');
    expect(docXml).not.toContain('w:headerReference');
    expect(docXml).not.toContain('w:footerReference');

    // Verify relationships file has no header/footer entries
    const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(relsXml).not.toContain('/header');
    expect(relsXml).not.toContain('/footer');
  });
});
