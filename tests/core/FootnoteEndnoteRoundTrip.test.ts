/**
 * Tests for footnote/endnote save pipeline integration
 * Covers: round-trip, passthrough, relationships, content types, clear API
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLParser } from '../../src/xml/XMLParser';

describe('Footnote/Endnote Save Pipeline', () => {
  describe('Programmatic creation', () => {
    it('should save footnotes created from scratch', async () => {
      const doc = Document.create();
      doc.createParagraph('Document text');
      const fn = doc.createFootnote('This is a footnote');
      expect(fn.getId()).toBeGreaterThan(0);

      const buffer = await doc.toBuffer();
      doc.dispose();

      // Verify footnotes.xml exists in the ZIP
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const footnotesXml = zip.getFileAsString('word/footnotes.xml');
      expect(footnotesXml).toBeDefined();
      expect(footnotesXml).toContain('w:footnotes');
      expect(footnotesXml).toContain('This is a footnote');
    });

    it('should save endnotes created from scratch', async () => {
      const doc = Document.create();
      doc.createParagraph('Document text');
      const en = doc.createEndnote('This is an endnote');
      expect(en.getId()).toBeGreaterThan(0);

      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const endnotesXml = zip.getFileAsString('word/endnotes.xml');
      expect(endnotesXml).toBeDefined();
      expect(endnotesXml).toContain('w:endnotes');
      expect(endnotesXml).toContain('This is an endnote');
    });

    it('should include footnotes relationship when footnotes are created', async () => {
      const doc = Document.create();
      doc.createParagraph('Text');
      doc.createFootnote('Note');

      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const relsXml = zip.getFileAsString('word/_rels/document.xml.rels');
      expect(relsXml).toBeDefined();

      const footnotesRelPattern = /Type="[^"]*\/footnotes"/g;
      const matches = relsXml!.match(footnotesRelPattern) || [];
      expect(matches).toHaveLength(1);
    });

    it('should include endnotes relationship when endnotes are created', async () => {
      const doc = Document.create();
      doc.createParagraph('Text');
      doc.createEndnote('Note');

      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const relsXml = zip.getFileAsString('word/_rels/document.xml.rels');
      expect(relsXml).toBeDefined();

      const endnotesRelPattern = /Type="[^"]*\/endnotes"/g;
      const matches = relsXml!.match(endnotesRelPattern) || [];
      expect(matches).toHaveLength(1);
    });

    it('should include footnotes in Content_Types', async () => {
      const doc = Document.create();
      doc.createParagraph('Text');
      doc.createFootnote('Note');

      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const contentTypes = zip.getFileAsString('[Content_Types].xml');
      expect(contentTypes).toBeDefined();
      expect(contentTypes).toContain('footnotes+xml');
    });

    it('should include endnotes in Content_Types', async () => {
      const doc = Document.create();
      doc.createParagraph('Text');
      doc.createEndnote('Note');

      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const contentTypes = zip.getFileAsString('[Content_Types].xml');
      expect(contentTypes).toBeDefined();
      expect(contentTypes).toContain('endnotes+xml');
    });
  });

  describe('Round-trip', () => {
    it('should preserve footnotes through round-trip', async () => {
      // Create document with footnotes
      const doc = Document.create();
      doc.createParagraph('Text with footnote');
      doc.createFootnote('First footnote');
      doc.createFootnote('Second footnote');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Round-trip
      const doc2 = await Document.loadFromBuffer(buffer1);
      expect(doc2.getFootnoteManager().getCount()).toBe(2);

      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Verify footnotes survived
      const doc3 = await Document.loadFromBuffer(buffer2);
      expect(doc3.getFootnoteManager().getCount()).toBe(2);
      doc3.dispose();
    });

    it('should preserve endnotes through round-trip', async () => {
      const doc = Document.create();
      doc.createParagraph('Text');
      doc.createEndnote('First endnote');
      doc.createEndnote('Second endnote');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buffer1);
      expect(doc2.getEndnoteManager().getCount()).toBe(2);

      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      const doc3 = await Document.loadFromBuffer(buffer2);
      expect(doc3.getEndnoteManager().getCount()).toBe(2);
      doc3.dispose();
    });

    it('should not create duplicate footnotes relationships on round-trip', async () => {
      const doc = Document.create();
      doc.createParagraph('Text');
      doc.createFootnote('Note');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Two round-trips
      const doc2 = await Document.loadFromBuffer(buffer1);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      const doc3 = await Document.loadFromBuffer(buffer2);
      const buffer3 = await doc3.toBuffer();
      doc3.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer3);
      const relsXml = zip.getFileAsString('word/_rels/document.xml.rels');
      const matches = relsXml!.match(/Type="[^"]*\/footnotes"/g) || [];
      expect(matches).toHaveLength(1);
    });

    it('should not create duplicate endnotes relationships on round-trip', async () => {
      const doc = Document.create();
      doc.createParagraph('Text');
      doc.createEndnote('Note');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buffer1);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      const doc3 = await Document.loadFromBuffer(buffer2);
      const buffer3 = await doc3.toBuffer();
      doc3.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer3);
      const relsXml = zip.getFileAsString('word/_rels/document.xml.rels');
      const matches = relsXml!.match(/Type="[^"]*\/endnotes"/g) || [];
      expect(matches).toHaveLength(1);
    });
  });

  describe('Passthrough', () => {
    it('should not generate footnotes.xml when no footnotes exist', async () => {
      const doc = Document.create();
      doc.createParagraph('Just text');
      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const footnotesXml = zip.getFileAsString('word/footnotes.xml');
      expect(footnotesXml).toBeUndefined();
    });

    it('should not generate endnotes.xml when no endnotes exist', async () => {
      const doc = Document.create();
      doc.createParagraph('Just text');
      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const endnotesXml = zip.getFileAsString('word/endnotes.xml');
      expect(endnotesXml).toBeUndefined();
    });
  });

  describe('Clear API', () => {
    it('should clear footnotes and regenerate with only separators', async () => {
      const doc = Document.create();
      doc.createParagraph('Text');
      doc.createFootnote('Note 1');
      doc.createFootnote('Note 2');
      expect(doc.getFootnoteManager().getCount()).toBe(2);

      doc.clearFootnotes();
      expect(doc.getFootnoteManager().getCount()).toBe(0);

      const buffer = await doc.toBuffer();
      doc.dispose();

      // File should still exist (with separators) because modified flag is set
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const footnotesXml = zip.getFileAsString('word/footnotes.xml');
      expect(footnotesXml).toBeDefined();
      // Should contain proper OOXML separator elements, not text runs
      expect(footnotesXml).toContain('w:footnotes');
      expect(footnotesXml).toContain('<w:separator/>');
      expect(footnotesXml).toContain('<w:continuationSeparator/>');
      expect(footnotesXml).not.toContain('Note 1');
      expect(footnotesXml).not.toContain('Note 2');
    });

    it('should clear endnotes and regenerate with only separators', async () => {
      const doc = Document.create();
      doc.createParagraph('Text');
      doc.createEndnote('Endnote 1');
      doc.createEndnote('Endnote 2');
      expect(doc.getEndnoteManager().getCount()).toBe(2);

      doc.clearEndnotes();
      expect(doc.getEndnoteManager().getCount()).toBe(0);

      const buffer = await doc.toBuffer();
      doc.dispose();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const endnotesXml = zip.getFileAsString('word/endnotes.xml');
      expect(endnotesXml).toBeDefined();
      // Should contain proper OOXML separator elements, not text runs
      expect(endnotesXml).toContain('w:endnotes');
      expect(endnotesXml).toContain('<w:separator/>');
      expect(endnotesXml).toContain('<w:continuationSeparator/>');
      expect(endnotesXml).not.toContain('Endnote 1');
      expect(endnotesXml).not.toContain('Endnote 2');
    });

    it('should clear footnotes from a loaded document', async () => {
      // Create doc with footnotes
      const doc = Document.create();
      doc.createParagraph('Text');
      doc.createFootnote('Original note');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Load and clear
      const doc2 = await Document.loadFromBuffer(buffer1);
      expect(doc2.getFootnoteManager().getCount()).toBe(1);
      doc2.clearFootnotes();
      expect(doc2.getFootnoteManager().getCount()).toBe(0);

      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Verify cleared
      const doc3 = await Document.loadFromBuffer(buffer2);
      expect(doc3.getFootnoteManager().getCount()).toBe(0);
      doc3.dispose();
    });

    it('should clear endnotes from a loaded document', async () => {
      const doc = Document.create();
      doc.createParagraph('Text');
      doc.createEndnote('Original endnote');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      const doc2 = await Document.loadFromBuffer(buffer1);
      expect(doc2.getEndnoteManager().getCount()).toBe(1);
      doc2.clearEndnotes();
      expect(doc2.getEndnoteManager().getCount()).toBe(0);

      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      const doc3 = await Document.loadFromBuffer(buffer2);
      expect(doc3.getEndnoteManager().getCount()).toBe(0);
      doc3.dispose();
    });
  });

  describe('Manager API', () => {
    it('should expose footnote manager', () => {
      const doc = Document.create();
      const mgr = doc.getFootnoteManager();
      expect(mgr).toBeDefined();
      expect(mgr.getCount()).toBe(0);
      doc.dispose();
    });

    it('should expose endnote manager', () => {
      const doc = Document.create();
      const mgr = doc.getEndnoteManager();
      expect(mgr).toBeDefined();
      expect(mgr.getCount()).toBe(0);
      doc.dispose();
    });

    it('should create footnotes with sequential IDs', () => {
      const doc = Document.create();
      const fn1 = doc.createFootnote('First');
      const fn2 = doc.createFootnote('Second');
      expect(fn1.getId()).toBe(1);
      expect(fn2.getId()).toBe(2);
      doc.dispose();
    });

    it('should create endnotes with sequential IDs', () => {
      const doc = Document.create();
      const en1 = doc.createEndnote('First');
      const en2 = doc.createEndnote('Second');
      expect(en1.getId()).toBe(1);
      expect(en2.getId()).toBe(2);
      doc.dispose();
    });
  });
});
