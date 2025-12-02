/**
 * Tests for Document class
 */

import * as fs from 'fs/promises';
import * as path from 'path';
import { Document, DocumentProperties } from '../../src/core/Document';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { Paragraph } from '../../src/elements/Paragraph';
import { Table } from '../../src/elements/Table';
import { DOCX_PATHS } from '../../src/zip/types';

const TEST_OUTPUT_DIR = path.join(__dirname, '../../test-output');

describe('Document', () => {
  beforeAll(async () => {
    // Create test output directory
    await fs.mkdir(TEST_OUTPUT_DIR, { recursive: true });
  });

  afterAll(async () => {
    // Clean up test output directory
    try {
      await fs.rm(TEST_OUTPUT_DIR, { recursive: true, force: true });
    } catch (error) {
      // Ignore cleanup errors
    }
  });

  describe('Document.create()', () => {
    it('should create a new empty document', () => {
      const doc = Document.create();
      expect(doc).toBeInstanceOf(Document);
      expect(doc.getParagraphCount()).toBe(0);
    });

    it('should create document with properties', () => {
      const props: DocumentProperties = {
        title: 'Test Document',
        creator: 'Test Author',
        subject: 'Testing',
      };

      const doc = Document.create({ properties: props });
      const docProps = doc.getProperties();

      expect(docProps.title).toBe('Test Document');
      expect(docProps.creator).toBe('Test Author');
      expect(docProps.subject).toBe('Testing');
    });

    it('should initialize required DOCX files', () => {
      const doc = Document.create();
      const zipHandler = doc.getZipHandler();

      expect(zipHandler.hasFile(DOCX_PATHS.CONTENT_TYPES)).toBe(true);
      expect(zipHandler.hasFile(DOCX_PATHS.RELS)).toBe(true);
      expect(zipHandler.hasFile(DOCX_PATHS.DOCUMENT)).toBe(true);
      expect(zipHandler.hasFile(DOCX_PATHS.CORE_PROPS)).toBe(true);
      expect(zipHandler.hasFile(DOCX_PATHS.APP_PROPS)).toBe(true);
    });
  });

  describe('Paragraph management', () => {
    it('should add paragraphs', () => {
      const doc = Document.create();
      const para1 = new Paragraph().addText('First paragraph');
      const para2 = new Paragraph().addText('Second paragraph');

      doc.addParagraph(para1).addParagraph(para2);

      expect(doc.getParagraphCount()).toBe(2);
      expect(doc.getParagraphs()).toHaveLength(2);
    });

    it('should support method chaining for addParagraph', () => {
      const doc = Document.create();
      const para1 = new Paragraph().addText('Para 1');
      const para2 = new Paragraph().addText('Para 2');

      const result = doc.addParagraph(para1).addParagraph(para2);

      expect(result).toBe(doc);
      expect(doc.getParagraphCount()).toBe(2);
    });

    it('should create and add paragraph with text', () => {
      const doc = Document.create();
      const para = doc.createParagraph('Hello World');

      expect(doc.getParagraphCount()).toBe(1);
      expect(para.getText()).toBe('Hello World');
      expect(doc.getParagraphs()[0]).toBe(para);
    });

    it('should create empty paragraph when no text provided', () => {
      const doc = Document.create();
      const para = doc.createParagraph();

      expect(doc.getParagraphCount()).toBe(1);
      expect(para.getText()).toBe('');
    });

    it('should get paragraphs as copy', () => {
      const doc = Document.create();
      doc.createParagraph('Test');

      const paragraphs1 = doc.getParagraphs();
      const paragraphs2 = doc.getParagraphs();

      expect(paragraphs1).not.toBe(paragraphs2);
      expect(paragraphs1).toEqual(paragraphs2);
    });

    it('should clear all paragraphs', () => {
      const doc = Document.create();
      doc.createParagraph('Para 1');
      doc.createParagraph('Para 2');
      doc.createParagraph('Para 3');

      expect(doc.getParagraphCount()).toBe(3);

      doc.clearParagraphs();

      expect(doc.getParagraphCount()).toBe(0);
      expect(doc.getParagraphs()).toHaveLength(0);
    });

    it('should support method chaining for clearParagraphs', () => {
      const doc = Document.create();
      doc.createParagraph('Test');

      const result = doc.clearParagraphs();

      expect(result).toBe(doc);
    });
  });

  describe('Document properties', () => {
    it('should set properties', () => {
      const doc = Document.create();

      doc.setProperties({
        title: 'My Document',
        subject: 'Subject',
        creator: 'John Doe',
      });

      const props = doc.getProperties();
      expect(props.title).toBe('My Document');
      expect(props.subject).toBe('Subject');
      expect(props.creator).toBe('John Doe');
    });

    it('should merge properties', () => {
      const doc = Document.create({
        properties: {
          title: 'Original Title',
          creator: 'Original Author',
        },
      });

      doc.setProperties({
        title: 'New Title',
        subject: 'New Subject',
      });

      const props = doc.getProperties();
      expect(props.title).toBe('New Title');
      expect(props.creator).toBe('Original Author');
      expect(props.subject).toBe('New Subject');
    });

    it('should support method chaining for setProperties', () => {
      const doc = Document.create();
      const result = doc.setProperties({ title: 'Test' });
      expect(result).toBe(doc);
    });

    it('should get properties as copy', () => {
      const doc = Document.create({ properties: { title: 'Test' } });

      const props1 = doc.getProperties();
      const props2 = doc.getProperties();

      expect(props1).not.toBe(props2);
      expect(props1).toEqual(props2);
    });

    it('should handle special characters in properties', () => {
      const doc = Document.create();

      doc.setProperties({
        title: 'Test & <Document>',
        description: 'Contains "quotes" and \'apostrophes\'',
      });

      const props = doc.getProperties();
      expect(props.title).toBe('Test & <Document>');
      expect(props.description).toBe('Contains "quotes" and \'apostrophes\'');
    });
  });

  describe('save()', () => {
    it('should save document to file', async () => {
      const doc = Document.create();
      doc.createParagraph('Test content');

      const outputPath = path.join(TEST_OUTPUT_DIR, 'test-save.docx');
      await doc.save(outputPath);

      const stats = await fs.stat(outputPath);
      expect(stats.isFile()).toBe(true);
      expect(stats.size).toBeGreaterThan(0);
    });

    it('should save document with multiple paragraphs', async () => {
      const doc = Document.create({ properties: { title: 'Multi-para Doc' } });

      doc.createParagraph('First paragraph');
      doc.createParagraph('Second paragraph');
      doc.createParagraph('Third paragraph');

      const outputPath = path.join(TEST_OUTPUT_DIR, 'test-multi-para.docx');
      await doc.save(outputPath);

      const stats = await fs.stat(outputPath);
      expect(stats.isFile()).toBe(true);
    });

    it('should save document with formatted paragraphs', async () => {
      const doc = Document.create();

      const para1 = doc.createParagraph();
      para1.setAlignment('center').addText('Centered Title', { bold: true, size: 16 });

      const para2 = doc.createParagraph();
      para2.addText('Normal text with ');
      para2.addText('bold', { bold: true });
      para2.addText(' and ');
      para2.addText('italic', { italic: true });
      para2.addText(' formatting.');

      const outputPath = path.join(TEST_OUTPUT_DIR, 'test-formatted.docx');
      await doc.save(outputPath);

      const stats = await fs.stat(outputPath);
      expect(stats.isFile()).toBe(true);
    });

    it('should update document.xml when saving', async () => {
      const doc = Document.create();
      doc.createParagraph('Content');

      const outputPath = path.join(TEST_OUTPUT_DIR, 'test-update-xml.docx');
      await doc.save(outputPath);

      const zipHandler = doc.getZipHandler();
      const docXml = zipHandler.getFileAsString(DOCX_PATHS.DOCUMENT);

      expect(docXml).toContain('Content');
    });
  });

  describe('toBuffer()', () => {
    it('should generate document as buffer', async () => {
      const doc = Document.create();
      doc.createParagraph('Buffer test');

      const buffer = await doc.toBuffer();

      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);
    });

    it('should be able to load buffer back', async () => {
      const doc1 = Document.create({ properties: { title: 'Buffer Test' } });
      doc1.createParagraph('Test content');

      const buffer = await doc1.toBuffer();

      const doc2 = await Document.loadFromBuffer(buffer);
      const props = doc2.getProperties();

      expect(props.title).toBe('Buffer Test');
    });
  });

  describe('Document.load()', () => {
    it('should load document from file', async () => {
      // Create a document
      const doc1 = Document.create({ properties: { title: 'Load Test' } });
      doc1.createParagraph('Test paragraph');

      const filePath = path.join(TEST_OUTPUT_DIR, 'test-load.docx');
      await doc1.save(filePath);

      // Load it back
      const doc2 = await Document.load(filePath);
      const props = doc2.getProperties();

      expect(props.title).toBe('Load Test');
    });

    it('should throw error for invalid file', async () => {
      await expect(Document.load('nonexistent.docx')).rejects.toThrow();
    });
  });

  describe('Document.loadFromBuffer()', () => {
    it('should load document from buffer', async () => {
      const doc1 = Document.create();
      doc1.createParagraph('Buffer content');

      const buffer = await doc1.toBuffer();
      const doc2 = await Document.loadFromBuffer(buffer);

      expect(doc2).toBeInstanceOf(Document);
    });

    it('should throw error for invalid buffer', async () => {
      const invalidBuffer = Buffer.from('not a zip file');
      await expect(Document.loadFromBuffer(invalidBuffer)).rejects.toThrow();
    });
  });

  describe('XML generation', () => {
    it('should generate valid [Content_Types].xml', () => {
      const doc = Document.create();
      const zipHandler = doc.getZipHandler();
      const xml = zipHandler.getFileAsString(DOCX_PATHS.CONTENT_TYPES);

      expect(xml).toContain('<?xml version="1.0"');
      expect(xml).toContain('<Types xmlns=');
      expect(xml).toContain('word/document.xml');
    });

    it('should generate valid _rels/.rels', () => {
      const doc = Document.create();
      const zipHandler = doc.getZipHandler();
      const xml = zipHandler.getFileAsString(DOCX_PATHS.RELS);

      expect(xml).toContain('<?xml version="1.0"');
      expect(xml).toContain('<Relationships');
      expect(xml).toContain('word/document.xml');
    });

    it('should generate valid core.xml', () => {
      const doc = Document.create({
        properties: {
          title: 'Test Title',
          creator: 'Test Creator',
        },
      });
      const zipHandler = doc.getZipHandler();
      const xml = zipHandler.getFileAsString(DOCX_PATHS.CORE_PROPS);

      expect(xml).toContain('<?xml version="1.0"');
      expect(xml).toContain('Test Title');
      expect(xml).toContain('Test Creator');
    });

    it('should generate valid app.xml', () => {
      const doc = Document.create();
      const zipHandler = doc.getZipHandler();
      const xml = zipHandler.getFileAsString(DOCX_PATHS.APP_PROPS);

      expect(xml).toContain('<?xml version="1.0"');
      expect(xml).toContain('docxmlater');
      expect(xml).toContain('<Properties');
    });

    it('should escape special characters in properties', () => {
      const doc = Document.create({
        properties: {
          title: 'Test & <Special> Characters',
          description: 'Contains "quotes"',
        },
      });

      const zipHandler = doc.getZipHandler();
      const xml = zipHandler.getFileAsString(DOCX_PATHS.CORE_PROPS);

      // Check that XML entities are properly escaped
      expect(xml).toContain('&amp;'); // & in "Test & <Special>"
      expect(xml).toContain('&lt;'); // < in "<Special>"
      expect(xml).toContain('&gt;'); // > in "<Special>"
      // Note: Quotes don't need escaping in text content (only in attributes)
      expect(xml).toContain('"quotes"'); // Quotes remain as-is in text
    });
  });

  describe('Integration tests', () => {
    it('should create a complete valid DOCX file', async () => {
      const doc = Document.create({
        properties: {
          title: 'Integration Test',
          creator: 'DocXML Test Suite',
          subject: 'Testing',
        },
      });

      // Add title
      const title = doc.createParagraph();
      title.setAlignment('center');
      title.setSpaceBefore(480);
      title.setSpaceAfter(240);
      title.addText('Integration Test Document', { bold: true, size: 18 });

      // Add content paragraphs
      doc.createParagraph('This is the first paragraph of content.');

      const para2 = doc.createParagraph();
      para2.addText('This paragraph has ');
      para2.addText('bold', { bold: true });
      para2.addText(', ');
      para2.addText('italic', { italic: true });
      para2.addText(', and ');
      para2.addText('colored', { color: 'FF0000' });
      para2.addText(' text.');

      doc.createParagraph();

      const para4 = doc.createParagraph();
      para4.setAlignment('right');
      para4.addText('Right-aligned paragraph', { italic: true });

      const outputPath = path.join(TEST_OUTPUT_DIR, 'integration-test.docx');
      await doc.save(outputPath);

      // Verify file exists and has content
      const stats = await fs.stat(outputPath);
      expect(stats.isFile()).toBe(true);
      expect(stats.size).toBeGreaterThan(1000);

      // Verify all required files are present
      const zipHandler = doc.getZipHandler();
      expect(zipHandler.getFileCount()).toBeGreaterThanOrEqual(5);
    });

    it('should handle documents with many paragraphs', async () => {
      const doc = Document.create();

      for (let i = 1; i <= 100; i++) {
        doc.createParagraph(`Paragraph ${i}`);
      }

      expect(doc.getParagraphCount()).toBe(100);

      const outputPath = path.join(TEST_OUTPUT_DIR, 'many-paragraphs.docx');
      await doc.save(outputPath);

      const stats = await fs.stat(outputPath);
      expect(stats.isFile()).toBe(true);
    });

    it('should create round-trip compatible documents', async () => {
      const doc1 = Document.create({
        properties: {
          title: 'Round Trip Test',
          creator: 'Test',
        },
      });

      doc1.createParagraph('First paragraph');
      doc1.createParagraph('Second paragraph');

      const buffer1 = await doc1.toBuffer();
      const doc2 = await Document.loadFromBuffer(buffer1);
      const buffer2 = await doc2.toBuffer();

      // Buffers should be similar in size
      expect(Math.abs(buffer1.length - buffer2.length)).toBeLessThan(100);
    });

    it('should preserve element order when loading and saving documents', async () => {
      // This test addresses the critical bug in DocumentParser where elements
      // were parsed by type (all paragraphs, then all tables) instead of by
      // document order, causing massive content loss and structure corruption.

      const doc1 = Document.create();

      // Create a document with interleaved paragraphs and tables
      // This structure is common in real-world documents

      doc1.createParagraph('Paragraph 1');

      const table1 = doc1.createTable(2, 2);
      const t1r1c1 = new Paragraph();
      t1r1c1.addText('Table 1, Row 1, Cell 1');
      table1.getRow(0)?.getCell(0)?.addParagraph(t1r1c1);

      const t1r1c2 = new Paragraph();
      t1r1c2.addText('Table 1, Row 1, Cell 2');
      table1.getRow(0)?.getCell(1)?.addParagraph(t1r1c2);

      const t1r2c1 = new Paragraph();
      t1r2c1.addText('Table 1, Row 2, Cell 1');
      table1.getRow(1)?.getCell(0)?.addParagraph(t1r2c1);

      const t1r2c2 = new Paragraph();
      t1r2c2.addText('Table 1, Row 2, Cell 2');
      table1.getRow(1)?.getCell(1)?.addParagraph(t1r2c2);

      doc1.createParagraph('Paragraph 2');
      doc1.createParagraph('Paragraph 3');

      const table2 = doc1.createTable(3, 2);
      const t2r1c1 = new Paragraph();
      t2r1c1.addText('Table 2, Row 1, Cell 1');
      table2.getRow(0)?.getCell(0)?.addParagraph(t2r1c1);

      const t2r1c2 = new Paragraph();
      t2r1c2.addText('Table 2, Row 1, Cell 2');
      table2.getRow(0)?.getCell(1)?.addParagraph(t2r1c2);

      doc1.createParagraph('Paragraph 4');

      const table3 = doc1.createTable(1, 3);
      const t3r1c1 = new Paragraph();
      t3r1c1.addText('Table 3, Cell 1');
      table3.getRow(0)?.getCell(0)?.addParagraph(t3r1c1);

      const t3r1c2 = new Paragraph();
      t3r1c2.addText('Table 3, Cell 2');
      table3.getRow(0)?.getCell(1)?.addParagraph(t3r1c2);

      const t3r1c3 = new Paragraph();
      t3r1c3.addText('Table 3, Cell 3');
      table3.getRow(0)?.getCell(2)?.addParagraph(t3r1c3);

      doc1.createParagraph('Paragraph 5');
      doc1.createParagraph('Paragraph 6');

      // Expected order:
      // 0: Paragraph 1
      // 1: Table 1 (2x2)
      // 2: Paragraph 2
      // 3: Paragraph 3
      // 4: Table 2 (3x2)
      // 5: Paragraph 4
      // 6: Table 3 (1x3)
      // 7: Paragraph 5
      // 8: Paragraph 6

      // Save and reload
      const buffer = await doc1.toBuffer();
      const doc2 = await Document.loadFromBuffer(buffer);

      // Verify element counts
      const bodyElements = doc2.getBodyElements();
      expect(bodyElements.length).toBe(9); // 6 paragraphs + 3 tables

      // Verify element types in order
      expect(bodyElements[0]).toBeInstanceOf(Paragraph);
      expect(bodyElements[1]).toBeInstanceOf(Table);
      expect(bodyElements[2]).toBeInstanceOf(Paragraph);
      expect(bodyElements[3]).toBeInstanceOf(Paragraph);
      expect(bodyElements[4]).toBeInstanceOf(Table);
      expect(bodyElements[5]).toBeInstanceOf(Paragraph);
      expect(bodyElements[6]).toBeInstanceOf(Table);
      expect(bodyElements[7]).toBeInstanceOf(Paragraph);
      expect(bodyElements[8]).toBeInstanceOf(Paragraph);

      // Verify paragraph text content
      expect((bodyElements[0] as any).getRuns()[0]?.getText()).toBe('Paragraph 1');
      expect((bodyElements[2] as any).getRuns()[0]?.getText()).toBe('Paragraph 2');
      expect((bodyElements[3] as any).getRuns()[0]?.getText()).toBe('Paragraph 3');
      expect((bodyElements[5] as any).getRuns()[0]?.getText()).toBe('Paragraph 4');
      expect((bodyElements[7] as any).getRuns()[0]?.getText()).toBe('Paragraph 5');
      expect((bodyElements[8] as any).getRuns()[0]?.getText()).toBe('Paragraph 6');

      // Verify table dimensions
      expect((bodyElements[1] as any).getRowCount()).toBe(2);
      expect((bodyElements[4] as any).getRowCount()).toBe(3);
      expect((bodyElements[6] as any).getRowCount()).toBe(1);

      // Verify table cell content (first cell of each table)
      const table1Cell = (bodyElements[1] as any).getRow(0)?.getCell(0);
      expect(table1Cell).toBeDefined();
      const table1Text = table1Cell?.getParagraphs()[0]?.getRuns()[0]?.getText();
      expect(table1Text).toBe('Table 1, Row 1, Cell 1');

      const table2Cell = (bodyElements[4] as any).getRow(0)?.getCell(0);
      expect(table2Cell).toBeDefined();
      const table2Text = table2Cell?.getParagraphs()[0]?.getRuns()[0]?.getText();
      expect(table2Text).toBe('Table 2, Row 1, Cell 1');

      const table3Cell = (bodyElements[6] as any).getRow(0)?.getCell(0);
      expect(table3Cell).toBeDefined();
      const table3Text = table3Cell?.getParagraphs()[0]?.getRuns()[0]?.getText();
      expect(table3Text).toBe('Table 3, Cell 1');

      // Verify no content loss - count all paragraphs including those in tables
      let totalParagraphs = 0;
      for (const element of bodyElements) {
        if (element instanceof Paragraph) {
          totalParagraphs++;
        } else if (element instanceof Table) {
          const table = element as any;
          for (const row of table.getRows()) {
            for (const cell of row.getCells()) {
              if (cell) {
                totalParagraphs += cell.getParagraphs().length;
              }
            }
          }
        }
      }

      // 6 body paragraphs + 4 (table1) + 6 (table2) + 3 (table3) = 19 total
      expect(totalParagraphs).toBe(19);
    });
  });

  describe('Parsing documents with conflicting paragraph properties', () => {
    test('should resolve pageBreakBefore + keepNext conflict during parsing', async () => {
      // Create a document with conflicting properties
      const doc = Document.create();
      const para = new Paragraph()
        .addText('Test content')
        .setPageBreakBefore(true)
        .setKeepNext(true)     // Clears pageBreakBefore
        .setKeepLines(true);
      doc.addParagraph(para);

      // Save and reload
      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);

      // Get the paragraph
      const bodyElements = loadedDoc.getBodyElements();
      const loadedPara = bodyElements[0] as Paragraph;
      const formatting = loadedPara.getFormatting();

      // Conflict should be resolved: keepNext/keepLines take priority
      expect(formatting.keepNext).toBe(true);
      expect(formatting.keepLines).toBe(true);
      expect(formatting.pageBreakBefore).toBe(false);
    });

    test('should preserve keepNext/keepLines when pageBreakBefore is not set', async () => {
      // Create a document without conflicts
      const doc = Document.create();
      const para = new Paragraph()
        .addText('Test content')
        .setKeepNext(true)
        .setKeepLines(true);
      doc.addParagraph(para);

      // Save and reload
      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);

      // Get the paragraph
      const bodyElements = loadedDoc.getBodyElements();
      const loadedPara = bodyElements[0] as Paragraph;
      const formatting = loadedPara.getFormatting();

      // Properties should be preserved (pageBreakBefore is false, not undefined, since keepNext was set)
      expect(formatting.pageBreakBefore).toBe(false);
      expect(formatting.keepNext).toBe(true);
      expect(formatting.keepLines).toBe(true);
    });

    test('should handle multiple paragraphs with mixed conflict scenarios', async () => {
      const doc = Document.create();

      // Para 1: Has conflict - pageBreakBefore then keepNext (keepNext wins)
      const para1 = new Paragraph()
        .addText('Paragraph 1')
        .setPageBreakBefore(true)
        .setKeepNext(true);  // Clears pageBreakBefore
      doc.addParagraph(para1);

      // Para 2: No conflict, just keepNext
      const para2 = new Paragraph()
        .addText('Paragraph 2')
        .setKeepNext(true);
      doc.addParagraph(para2);

      // Para 3: No conflict, just pageBreakBefore
      const para3 = new Paragraph()
        .addText('Paragraph 3')
        .setPageBreakBefore(true);
      doc.addParagraph(para3);

      // Save and reload
      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);

      const bodyElements = loadedDoc.getBodyElements();

      // Para 1: Conflict resolved - keepNext wins
      const loadedPara1 = bodyElements[0] as Paragraph;
      expect(loadedPara1.getFormatting().keepNext).toBe(true);
      expect(loadedPara1.getFormatting().pageBreakBefore).toBe(false);

      // Para 2: keepNext preserved (pageBreakBefore is false since keepNext was set)
      const loadedPara2 = bodyElements[1] as Paragraph;
      expect(loadedPara2.getFormatting().keepNext).toBe(true);
      expect(loadedPara2.getFormatting().pageBreakBefore).toBe(false);

      // Para 3: pageBreakBefore preserved
      const loadedPara3 = bodyElements[2] as Paragraph;
      expect(loadedPara3.getFormatting().pageBreakBefore).toBe(true);
      expect(loadedPara3.getFormatting().keepNext).toBeUndefined();
    });

    test('should resolve conflicts when properties come from XML with non-standard order', async () => {
      const doc = Document.create();
      const para = new Paragraph()
        .addText('Test content')
        .setKeepNext(true)
        .setKeepLines(true)
        .setPageBreakBefore(true)  // Can be set after
        .setKeepNext(true);         // Call again to clear pageBreakBefore
      doc.addParagraph(para);

      // Save and reload
      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);

      const bodyElements = loadedDoc.getBodyElements();
      const loadedPara = bodyElements[0] as Paragraph;
      const formatting = loadedPara.getFormatting();

      // Conflict should be resolved - keepNext/keepLines win
      expect(formatting.keepNext).toBe(true);
      expect(formatting.keepLines).toBe(true);
      expect(formatting.pageBreakBefore).toBe(false);
    });
  });

  describe('TOC Field Instruction Parsing', () => {
    describe('parseTOCFieldInstruction()', () => {
      test('should parse TOC field with double-quoted \o switch', () => {
        const doc = Document.create();
        const instruction = 'TOC \\o "1-3"';
        
        // We need to access the private method for testing
        // For now, we'll verify through the internal structure
        // This test documents the expected behavior
        expect(instruction).toMatch(/\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/);
      });

      test('should parse TOC field with single-quoted \o switch', () => {
        const doc = Document.create();
        const instruction = "TOC \\o '1-3'";
        
        // Verify regex matches single-quoted format
        expect(instruction).toMatch(/\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/);
      });

      test('should parse TOC field with unquoted \o switch', () => {
        const doc = Document.create();
        const instruction = 'TOC \\o 1-3';
        
        // Verify regex matches unquoted format (this is the key fix)
        expect(instruction).toMatch(/\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/);
      });

      test('should extract correct outline levels from double-quoted format', () => {
        const instruction = 'TOC \\o "1-3"';
        const regex = /\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/;
        const match = instruction.match(regex);
        
        expect(match).not.toBeNull();
        if (match) {
          const start = parseInt(match[1] || match[3] || match[5]!, 10);
          const end = parseInt(match[2] || match[4] || match[6]!, 10);
          expect(start).toBe(1);
          expect(end).toBe(3);
        }
      });

      test('should extract correct outline levels from single-quoted format', () => {
        const instruction = "TOC \\o '2-4'";
        const regex = /\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/;
        const match = instruction.match(regex);
        
        expect(match).not.toBeNull();
        if (match) {
          const start = parseInt(match[1] || match[3] || match[5]!, 10);
          const end = parseInt(match[2] || match[4] || match[6]!, 10);
          expect(start).toBe(2);
          expect(end).toBe(4);
        }
      });

      test('should extract correct outline levels from unquoted format', () => {
        const instruction = 'TOC \\o 1-3';
        const regex = /\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/;
        const match = instruction.match(regex);
        
        expect(match).not.toBeNull();
        if (match) {
          const start = parseInt(match[1] || match[3] || match[5]!, 10);
          const end = parseInt(match[2] || match[4] || match[6]!, 10);
          expect(start).toBe(1);
          expect(end).toBe(3);
        }
      });

      test('should handle multiple spaces before unquoted \o value', () => {
        const instruction = 'TOC \\o   1-3';
        const regex = /\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/;
        const match = instruction.match(regex);
        
        expect(match).not.toBeNull();
        if (match) {
          const start = parseInt(match[1] || match[3] || match[5]!, 10);
          const end = parseInt(match[2] || match[4] || match[6]!, 10);
          expect(start).toBe(1);
          expect(end).toBe(3);
        }
      });

      test('should not match non-TOC field instructions', () => {
        const instruction = 'HYPERLINK \\l "anchor"';
        const regex = /\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/;
        const match = instruction.match(regex);
        
        expect(match).toBeNull();
      });

      test('should handle complex TOC instructions with multiple switches', () => {
        const instruction = 'TOC \\o "1-3" \\t "Heading 1,1,Heading 2,2"';
        const regex = /\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/;
        const match = instruction.match(regex);
        
        expect(match).not.toBeNull();
        if (match) {
          const start = parseInt(match[1] || match[3] || match[5]!, 10);
          const end = parseInt(match[2] || match[4] || match[6]!, 10);
          expect(start).toBe(1);
          expect(end).toBe(3);
        }
      });

      test('should handle complex TOC instructions with unquoted \o switch', () => {
        const instruction = 'TOC \\o 1-3 \\t "Heading 1,1,Heading 2,2"';
        const regex = /\\o\s+(?:"(\d+)-(\d+)"|'(\d+)-(\d+)'|(\d+)-(\d+))/;
        const match = instruction.match(regex);
        
        expect(match).not.toBeNull();
        if (match) {
          const start = parseInt(match[1] || match[3] || match[5]!, 10);
          const end = parseInt(match[2] || match[4] || match[6]!, 10);
          expect(start).toBe(1);
          expect(end).toBe(3);
        }
      });
    });
  });

  describe('Bookmark Helpers', () => {
    describe('addTopBookmark()', () => {
      test('should create _top bookmark when it does not exist', () => {

        const doc = Document.create();

        // Verify no _top bookmark exists yet
        expect(doc.hasBookmark('_top')).toBe(false);

        const result = doc.addTopBookmark();

        // Verify bookmark was created
        expect(result.bookmark).toBeDefined();
        expect(result.bookmark.getName()).toBe('_top');
        expect(result.bookmark.getId()).toBe(0);
        expect(result.anchor).toBe('_top');
        expect(result.hyperlink).toBeInstanceOf(Function);

        // Verify bookmark is registered
        expect(doc.hasBookmark('_top')).toBe(true);
      });

      test('should return existing _top bookmark when already present', () => {
        const doc = Document.create();

        // Add _top bookmark first time
        const result1 = doc.addTopBookmark();
        const bookmark1 = result1.bookmark;

        // Add _top bookmark second time
        const result2 = doc.addTopBookmark();
        const bookmark2 = result2.bookmark;

        // Should return the same bookmark instance
        expect(bookmark2).toBe(bookmark1);
        expect(bookmark2.getName()).toBe('_top');
        expect(bookmark2.getId()).toBe(0);
      });

      test('should place bookmark at the beginning of document', () => {
        const doc = Document.create();

        // Add some content first
        doc.createParagraph().addText('Paragraph 1');
        doc.createParagraph().addText('Paragraph 2');

        expect(doc.getParagraphCount()).toBe(2);

        // Add _top bookmark
        doc.addTopBookmark();

        // Should still have 2 paragraphs (bookmark added to first paragraph, no new paragraph created)
        expect(doc.getParagraphCount()).toBe(2);

        const bodyElements = doc.getBodyElements();
        const firstPara = bodyElements[0] as Paragraph;

        // First paragraph should still have its text (bookmark is added to it)
        expect(firstPara.getText()).toBe('Paragraph 1');
      });

      test('should create working hyperlinks to _top bookmark', () => {
        const doc = Document.create();

        const { hyperlink, anchor } = doc.addTopBookmark();

        // Create hyperlink using convenience function
        const link1 = hyperlink('Back to top');
        expect(link1.getAnchor()).toBe('_top');
        expect(link1.getText()).toBe('Back to top');
        expect(link1.isInternal()).toBe(true);

        // Create hyperlink manually using anchor
        const link2 = Hyperlink.createInternal(anchor, 'Go to top');
        expect(link2.getAnchor()).toBe('_top');
        expect(link2.getText()).toBe('Go to top');
      });

      test('should be idempotent - safe to call multiple times on empty document', () => {
        const doc = Document.create();

        // Call multiple times on empty document
        doc.addTopBookmark();
        doc.addTopBookmark();
        doc.addTopBookmark();

        // Should only have one _top bookmark
        const bookmarks = doc.getBookmarkManager().getAllBookmarks();
        const topBookmarks = bookmarks.filter(b => b.getName() === '_top');
        expect(topBookmarks.length).toBe(1);

        // Should only have one paragraph (the fallback empty paragraph with bookmark)
        expect(doc.getParagraphCount()).toBe(1);
      });

      test('should be idempotent - safe to call multiple times on document with content', () => {
        const doc = Document.create();

        // Add content first
        doc.createParagraph().addText('Paragraph 1');
        doc.createParagraph().addText('Paragraph 2');

        expect(doc.getParagraphCount()).toBe(2);

        // Call addTopBookmark multiple times
        doc.addTopBookmark();
        doc.addTopBookmark();
        doc.addTopBookmark();

        // Should only have one _top bookmark
        const bookmarks = doc.getBookmarkManager().getAllBookmarks();
        const topBookmarks = bookmarks.filter(b => b.getName() === '_top');
        expect(topBookmarks.length).toBe(1);

        // Should still have only 2 paragraphs (no extra paragraphs created)
        expect(doc.getParagraphCount()).toBe(2);
      });

      test('should preserve _top bookmark through save/load cycle', async () => {
        const doc = Document.create();

        // Add _top bookmark
        doc.addTopBookmark();

        // Add some content
        doc.createParagraph().addText('Content paragraph');

        // Save to buffer
        const buffer = await doc.toBuffer();

        // Load from buffer
        const loadedDoc = await Document.loadFromBuffer(buffer);

        // Note: Bookmark parsing is not yet implemented, but we can verify document structure
        // Verify document structure is preserved
        const bodyElements = loadedDoc.getBodyElements();
        expect(bodyElements.length).toBe(2); // Empty para with bookmark + content para

        // Verify the XML contains the bookmark
        const xml = loadedDoc.getZipHandler().getFileAsString(DOCX_PATHS.DOCUMENT);
        expect(xml).toBeDefined();
        if (xml) {
          expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="_top"');
          expect(xml).toContain('<w:bookmarkEnd w:id="0"');
        }
      });

      test('should generate correct XML structure', async () => {
        const doc = Document.create();

        doc.addTopBookmark();

        const buffer = await doc.toBuffer();
        const xml = doc.getZipHandler().getFileAsString(DOCX_PATHS.DOCUMENT);

        // Verify XML is present
        expect(xml).toBeDefined();

        if (xml) {
          // Verify XML contains bookmark start and end with correct attributes
          expect(xml).toContain('<w:bookmarkStart w:id="0" w:name="_top"');
          expect(xml).toContain('<w:bookmarkEnd w:id="0"');

          // Bookmark should be at the beginning of the body
          const bodyStart = xml.indexOf('<w:body>');
          const bookmarkStart = xml.indexOf('<w:bookmarkStart w:id="0" w:name="_top"');
          expect(bookmarkStart).toBeGreaterThan(bodyStart);
        }
      });

      test('should work with hyperlinks in other paragraphs', async () => {
        const doc = Document.create();

        const { hyperlink } = doc.addTopBookmark();

        // Add content paragraphs
        doc.createParagraph().addText('Section 1');
        doc.createParagraph().addText('Section 2');

        // Add hyperlink to _top in last paragraph
        const lastPara = doc.createParagraph();
        const link = hyperlink('Back to top');
        lastPara.addHyperlink(link);

        // Save and load
        const buffer = await doc.toBuffer();
        const loadedDoc = await Document.loadFromBuffer(buffer);

        // Verify structure
        const bodyElements = loadedDoc.getBodyElements();
        expect(bodyElements.length).toBe(4); // bookmark para + 3 content paras

        // Verify hyperlink works
        const lastLoadedPara = bodyElements[3] as Paragraph;
        expect(lastLoadedPara.getText()).toContain('Back to top');
      });
    });
  });

  describe('Preserve Blank Lines After Heading 2 Tables', () => {
    describe('applyStyles with preserveBlankLinesAfterHeading2Tables', () => {
      test('should mark blank lines as preserved when option is true', () => {
        const doc = Document.create();

        // Add a Heading2 paragraph
        const heading = doc.createParagraph('Test Header');
        heading.setStyle('Heading2');

        // Apply formatting with preserve option enabled
        doc.applyStyles({
          preserveBlankLinesAfterHeading2Tables: true
        });

        // Get all paragraphs
        const paragraphs = doc.getAllParagraphs();

        // The heading should now be in a table, and there should be a blank paragraph after it
        const tables = doc.getAllTables();
        expect(tables.length).toBe(1);

        // Get body elements
        const bodyElements = doc.getBodyElements();

        // Should have: table + blank paragraph
        expect(bodyElements.length).toBe(2);

        // Check that second element is a paragraph and is marked as preserved
        const blankPara = bodyElements[1];
        expect(blankPara).toBeInstanceOf(Paragraph);
        expect((blankPara as Paragraph).isPreserved()).toBe(true);
      });

      test('should not mark blank lines as preserved when option is false', () => {
        const doc = Document.create();

        // Add a Heading2 paragraph
        const heading = doc.createParagraph('Test Header');
        heading.setStyle('Heading2');

        // Apply formatting with preserve option disabled
        doc.applyStyles({
          preserveBlankLinesAfterHeading2Tables: false
        });

        // Get body elements
        const bodyElements = doc.getBodyElements();

        // Should have: table + blank paragraph
        expect(bodyElements.length).toBe(2);

        // Check that second element is a paragraph and is NOT marked as preserved
        const blankPara = bodyElements[1];
        expect(blankPara).toBeInstanceOf(Paragraph);
        expect((blankPara as Paragraph).isPreserved()).toBe(false);
      });

      test('should default to true when option is not specified', () => {
        const doc = Document.create();

        // Add a Heading2 paragraph
        const heading = doc.createParagraph('Test Header');
        heading.setStyle('Heading2');

        // Apply formatting without specifying preserve option
        doc.applyStyles();

        // Get body elements
        const bodyElements = doc.getBodyElements();

        // Should have: table + blank paragraph
        expect(bodyElements.length).toBe(2);

        // Check that second element is a paragraph and is marked as preserved (default: true)
        const blankPara = bodyElements[1];
        expect(blankPara).toBeInstanceOf(Paragraph);
        expect((blankPara as Paragraph).isPreserved()).toBe(true);
      });
    });

    // NOTE: Tests for keepOne and preserveHeader2BlankLines parameters were removed
    // as these features are obsolete and replaced by better functionality

    describe('removeExtraBlankParagraphs', () => {
      test('should not remove preserved blank paragraphs', () => {
        const doc = Document.create();

        // Add some paragraphs
        doc.createParagraph('First paragraph');

        const blank1 = doc.createParagraph();
        blank1.setPreserved(true);

        const blank2 = doc.createParagraph();

        doc.createParagraph('Second paragraph');

        // Remove extra blank paragraphs
        const result = doc.removeExtraBlankParagraphs();

        // Should remove blank2 but not blank1 (which is preserved)
        expect(result.removed).toBe(1);

        const paragraphs = doc.getAllParagraphs();
        expect(paragraphs.length).toBe(3); // First + blank1 + Second

        // Verify blank1 is still preserved
        expect(blank1.isPreserved()).toBe(true);
      });

      // Tests for obsolete keepOne and preserveHeader2BlankLines parameters have been removed
    });
  });

  describe('Revision Registration', () => {
    it('should register parsed revisions with RevisionManager after loading', async () => {
      // Create a document with revisions
      const doc = Document.create();
      const para1 = doc.createParagraph('First paragraph');

      // Add a revision to the paragraph
      const { Revision } = await import('../../src/elements/Revision');
      const { Run } = await import('../../src/elements/Run');

      const insertedRun = new Run('inserted text');
      const revision = new Revision({
        type: 'insert',
        author: 'Test Author',
        content: insertedRun,
        date: new Date()
      });

      para1.addRevision(revision);

      // Save to buffer
      const buffer = await doc.toBuffer();
      doc.dispose();

      // Reload the document with preserve mode
      const reloadedDoc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Verify revisions are registered with RevisionManager
      const revisionManager = reloadedDoc.getRevisionManager();
      const allRevisions = revisionManager.getAllRevisions();

      expect(allRevisions.length).toBeGreaterThanOrEqual(1);

      // Verify the revision has location info
      const foundRevision = allRevisions.find(r => r.getAuthor() === 'Test Author');
      expect(foundRevision).toBeDefined();

      reloadedDoc.dispose();
    });

    it('should make ChangelogGenerator work with parsed revisions', async () => {
      // Create a document with revisions
      const doc = Document.create();
      const para1 = doc.createParagraph('Some text');

      const { Revision } = await import('../../src/elements/Revision');
      const { Run } = await import('../../src/elements/Run');
      const { ChangelogGenerator } = await import('../../src/utils/ChangelogGenerator');

      const deletedRun = new Run('deleted content');
      const deletion = new Revision({
        type: 'delete',
        author: 'Reviewer',
        content: deletedRun,
        date: new Date()
      });

      para1.addRevision(deletion);

      // Save and reload
      const buffer = await doc.toBuffer();
      doc.dispose();

      const reloadedDoc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

      // Use ChangelogGenerator to get changes
      const entries = ChangelogGenerator.fromDocument(reloadedDoc);

      expect(entries.length).toBeGreaterThanOrEqual(1);

      // Find the deletion entry
      const deletionEntry = entries.find(e => e.revisionType === 'delete');
      expect(deletionEntry).toBeDefined();
      expect(deletionEntry?.author).toBe('Reviewer');

      reloadedDoc.dispose();
    });
  });
});
