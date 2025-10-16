/**
 * Tests for Document class
 */

import { Document, DocumentProperties } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { DOCX_PATHS } from '../../src/zip/types';
import * as fs from 'fs/promises';
import * as path from 'path';

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
      expect(xml).toContain('DocXML');
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
  });
});
