/**
 * Regression tests for DOCX corruption fixes (v9.5.37)
 *
 * Tests for 4 bugs that caused Word to report "unreadable content" errors:
 * 1. Missing mc:Ignorable on <w:document> root element
 * 2. Orphaned numId references pointing to removed numbering definitions
 * 3. Missing tracked change authors in people.xml
 * 4. Wrong attribute order in pPrChange elements
 */

import * as fs from 'fs/promises';
import * as path from 'path';
import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Revision } from '../../src/elements/Revision';
import { Run } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { AbstractNumbering } from '../../src/formatting/AbstractNumbering';
import { NumberingInstance } from '../../src/formatting/NumberingInstance';
import { NumberingLevel } from '../../src/formatting/NumberingLevel';

const TEST_OUTPUT_DIR = path.join(__dirname, '../../test-output');

describe('DOCX Corruption Fixes (v9.5.37)', () => {
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

  describe('Bug 1: mc:Ignorable preservation', () => {
    it('should auto-generate mc:Ignorable when extended namespaces are declared', () => {
      // XMLBuilder.createDocument() should add mc:Ignorable for w14/w15/wp14
      const xml = XMLBuilder.createDocument([]);

      expect(xml).toContain('mc:Ignorable=');
      expect(xml).toContain('w14');
      expect(xml).toContain('w15');
      expect(xml).toContain('wp14');
    });

    it('should preserve mc:Ignorable from loaded namespaces', () => {
      const namespaces: Record<string, string> = {
        'mc:Ignorable': 'w14 w15 w16se w16cid',
        'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
      };

      const xml = XMLBuilder.createDocument([], namespaces);

      // Should preserve the exact value from loaded namespaces, not regenerate
      expect(xml).toContain('mc:Ignorable="w14 w15 w16se w16cid"');
    });

    it('should not duplicate mc:Ignorable if already present', () => {
      const namespaces: Record<string, string> = {
        'mc:Ignorable': 'w14 w15',
        'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
      };

      const xml = XMLBuilder.createDocument([], namespaces);

      // Count occurrences - should only appear once
      const matches = xml.match(/mc:Ignorable=/g);
      expect(matches).toHaveLength(1);
    });

    it('should include w14, w15, and wp14 in generated mc:Ignorable', () => {
      const xml = XMLBuilder.createDocument([]);

      // Extract mc:Ignorable value
      const match = xml.match(/mc:Ignorable="([^"]+)"/);
      expect(match).not.toBeNull();
      expect(match).toBeDefined();

      const ignorableValues = match![1]!.split(' ');
      expect(ignorableValues).toContain('w14');
      expect(ignorableValues).toContain('w15');
      expect(ignorableValues).toContain('wp14');
    });

    it('should round-trip mc:Ignorable through load and save', async () => {
      const doc = Document.create();
      doc.createParagraph('Test');

      const buffer1 = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buffer1);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Extract document.xml and check mc:Ignorable
      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(buffer2);
      const docXml = await zip.file('word/document.xml')?.async('string');

      expect(docXml).toContain('mc:Ignorable=');
    });
  });

  describe('Bug 2: Orphaned numId reference validation', () => {
    it('should remove numbering from paragraphs referencing non-existent numId', () => {
      const doc = Document.create();
      const para = doc.createParagraph('Test list item');

      // Manually set a numbering reference to a numId that doesn't exist
      para.setNumbering(9999, 0);
      expect(para.getNumbering()).toBeTruthy();
      expect(para.getNumbering()?.numId).toBe(9999);

      // Validate should detect and fix the orphaned reference
      const fixed = doc.validateNumberingReferences();
      expect(fixed).toBe(1);

      // The paragraph should no longer have numbering
      expect(para.getNumbering()).toBeUndefined();

      doc.dispose();
    });

    it('should not remove numbering for valid numId references', () => {
      const doc = Document.create();

      // Create a list definition using the correct API
      const numberingManager = doc.getNumberingManager();
      const abstractNum = new AbstractNumbering(100);
      abstractNum.addLevel(NumberingLevel.createDecimalLevel(0));
      numberingManager.addAbstractNumbering(abstractNum);

      const numInstance = new NumberingInstance(100, 100);
      numberingManager.addNumberingInstance(numInstance);

      // Create paragraph with valid numbering
      const para = doc.createParagraph('Valid list item');
      para.setNumbering(100, 0);

      const fixed = doc.validateNumberingReferences();
      expect(fixed).toBe(0);

      // Numbering should still be present
      expect(para.getNumbering()).toBeTruthy();
      expect(para.getNumbering()?.numId).toBe(100);

      doc.dispose();
    });

    it('should run numId validation during save pipeline', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Orphaned list item');
      para.setNumbering(8888, 0);

      // Save should automatically validate and fix
      const buffer = await doc.toBuffer();
      doc.dispose();

      // Load back and verify no numbering on the paragraph
      const doc2 = await Document.loadFromBuffer(buffer);
      const paras = doc2.getParagraphs();
      expect(paras.length).toBeGreaterThan(0);

      // The paragraph should have had its numbering removed
      const firstPara = paras[0];
      expect(firstPara).toBeDefined();
      const numbering = firstPara!.getNumbering();
      expect(numbering).toBeUndefined();

      doc2.dispose();
    });
  });

  describe('Bug 3: people.xml author synchronization', () => {
    it('should add missing tracked change authors to people.xml', async () => {
      const doc = Document.create();

      // Add a paragraph with a tracked change revision
      const para = doc.createParagraph();
      const insertion = Revision.createInsertion('TestAuthor', new Run('inserted text'));
      para.addRevision(insertion);

      const buffer = await doc.toBuffer();
      doc.dispose();

      // Check people.xml in the output
      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(buffer);
      const peopleXml = await zip.file('word/people.xml')?.async('string');

      expect(peopleXml).toBeTruthy();
      expect(peopleXml).toContain('TestAuthor');
      expect(peopleXml).toContain('w15:person');
      expect(peopleXml).toContain('w15:presenceInfo');
    });

    it('should handle multiple tracked change authors', async () => {
      const doc = Document.create();

      const para1 = doc.createParagraph();
      para1.addRevision(Revision.createInsertion('Author1', new Run('text1')));

      const para2 = doc.createParagraph();
      para2.addRevision(Revision.createDeletion('Author2', new Run('text2')));

      const buffer = await doc.toBuffer();
      doc.dispose();

      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(buffer);
      const peopleXml = await zip.file('word/people.xml')?.async('string');

      expect(peopleXml).toContain('Author1');
      expect(peopleXml).toContain('Author2');
    });

    it('should collect authors from pPrChange', async () => {
      const doc = Document.create();

      const para = doc.createParagraph('test');
      para.setParagraphPropertiesChange({
        author: 'PPrChangeAuthor',
        date: '2025-01-01T00:00:00Z',
        id: '1',
      });

      const buffer = await doc.toBuffer();
      doc.dispose();

      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(buffer);
      const peopleXml = await zip.file('word/people.xml')?.async('string');

      expect(peopleXml).toBeTruthy();
      expect(peopleXml).toContain('PPrChangeAuthor');
    });

    it('should not create people.xml when no tracked changes exist', async () => {
      const doc = Document.create();
      doc.createParagraph('No revisions here');

      const buffer = await doc.toBuffer();
      doc.dispose();

      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(buffer);
      const peopleXml = await zip.file('word/people.xml')?.async('string');

      // No people.xml should exist when there are no tracked changes
      expect(peopleXml).toBeUndefined();
    });

    it('should not duplicate authors already present in people.xml', async () => {
      const doc = Document.create();

      // Create two revisions by the same author
      const para1 = doc.createParagraph();
      para1.addRevision(Revision.createInsertion('SameAuthor', new Run('text1')));

      const para2 = doc.createParagraph();
      para2.addRevision(Revision.createInsertion('SameAuthor', new Run('text2')));

      const buffer = await doc.toBuffer();
      doc.dispose();

      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(buffer);
      const peopleXml = await zip.file('word/people.xml')?.async('string');

      // Should only have one author entry for SameAuthor
      const authorMatches = peopleXml?.match(/w15:author="SameAuthor"/g);
      expect(authorMatches).toHaveLength(1);
    });
  });

  describe('Bug 4: pPrChange attribute order', () => {
    it('should serialize pPrChange with w:id before w:author before w:date', () => {
      const para = new Paragraph();
      para.setParagraphPropertiesChange({
        id: '42',
        author: 'TestAuthor',
        date: '2025-01-15T10:00:00Z',
        previousProperties: {
          alignment: 'left',
        },
      });

      const xml = para.toXML();
      // Convert XMLElement to string for inspection
      const xmlStr = XMLBuilder.elementToString(xml);

      // Verify w:pPrChange has correct attribute order: w:id first, then w:author, then w:date
      const pPrChangeMatch = xmlStr.match(/w:pPrChange([^>]*)/);
      expect(pPrChangeMatch).not.toBeNull();

      const attrs = pPrChangeMatch![1]!;
      const idPos = attrs.indexOf('w:id=');
      const authorPos = attrs.indexOf('w:author=');
      const datePos = attrs.indexOf('w:date=');

      expect(idPos).toBeGreaterThan(-1);
      expect(authorPos).toBeGreaterThan(-1);
      expect(datePos).toBeGreaterThan(-1);

      // w:id should come before w:author, which should come before w:date
      expect(idPos).toBeLessThan(authorPos);
      expect(authorPos).toBeLessThan(datePos);
    });
  });
});
