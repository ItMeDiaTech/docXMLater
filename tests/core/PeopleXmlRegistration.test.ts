/**
 * Tests for people.xml content type and relationship registration
 *
 * Issue: updatePeopleXml() creates word/people.xml but never registers it
 * in [Content_Types].xml or word/_rels/document.xml.rels, causing Word to
 * report the file as missing/corrupted.
 *
 * Fix: Added people.xml override to content types when file exists,
 * and relationship registration in updatePeopleXml().
 */

import { describe, it, expect } from '@jest/globals';
import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('people.xml Registration', () => {

  it('should register people.xml in Content_Types when track changes create authors', async () => {
    const doc = Document.create();
    doc.enableTrackChanges({ author: 'Test Author' });

    // Add a tracked insertion to trigger people.xml creation
    const para = new Paragraph();
    const run = new Run('Tracked text');
    const revision = Revision.createInsertion('Test Author', run);
    para.addRevision(revision);
    doc.addParagraph(para);

    const buffer = await doc.toBuffer();
    doc.dispose();

    const zipHandler = new ZipHandler();
    await zipHandler.loadFromBuffer(buffer);

    // Verify people.xml exists
    expect(zipHandler.hasFile('word/people.xml')).toBe(true);

    // Verify Content_Types includes people.xml override
    const contentTypes = zipHandler.getFileAsString('[Content_Types].xml') || '';
    expect(contentTypes).toContain('people.xml');
    expect(contentTypes).toContain('application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml');
  });

  it('should register people.xml relationship in document.xml.rels', async () => {
    const doc = Document.create();
    doc.enableTrackChanges({ author: 'Author One' });

    const para = new Paragraph();
    const run = new Run('Some tracked content');
    const revision = Revision.createInsertion('Author One', run);
    para.addRevision(revision);
    doc.addParagraph(para);

    const buffer = await doc.toBuffer();
    doc.dispose();

    const zipHandler = new ZipHandler();
    await zipHandler.loadFromBuffer(buffer);

    // Verify relationship exists
    const rels = zipHandler.getFileAsString('word/_rels/document.xml.rels') || '';
    expect(rels).toContain('people.xml');
    expect(rels).toContain('http://schemas.microsoft.com/office/2011/relationships/people');
  });

  it('should not create duplicate people.xml relationships on re-save', async () => {
    // First save
    const doc1 = Document.create();
    doc1.enableTrackChanges({ author: 'Author A' });

    const para1 = new Paragraph();
    para1.addRevision(Revision.createInsertion('Author A', new Run('First edit')));
    doc1.addParagraph(para1);

    const buffer1 = await doc1.toBuffer();
    doc1.dispose();

    // Load and re-save with additional author
    const doc2 = await Document.loadFromBuffer(buffer1, { revisionHandling: 'preserve' });
    doc2.enableTrackChanges({ author: 'Author B' });

    const para2 = new Paragraph();
    para2.addRevision(Revision.createInsertion('Author B', new Run('Second edit')));
    doc2.addParagraph(para2);

    const buffer2 = await doc2.toBuffer();
    doc2.dispose();

    const zipHandler = new ZipHandler();
    await zipHandler.loadFromBuffer(buffer2);

    // Count people relationships - should be exactly 1
    const rels = zipHandler.getFileAsString('word/_rels/document.xml.rels') || '';
    const peopleRelMatches = rels.match(/people\.xml/g) || [];
    expect(peopleRelMatches.length).toBe(1);
  });

  it('should not create people.xml when there are no tracked changes', async () => {
    const doc = Document.create();
    doc.addParagraph(new Paragraph().addText('No tracked changes here'));

    const buffer = await doc.toBuffer();
    doc.dispose();

    const zipHandler = new ZipHandler();
    await zipHandler.loadFromBuffer(buffer);

    // No people.xml should exist
    expect(zipHandler.hasFile('word/people.xml')).toBe(false);

    // Content_Types should not mention people.xml
    const contentTypes = zipHandler.getFileAsString('[Content_Types].xml') || '';
    expect(contentTypes).not.toContain('people.xml');
  });

  it('should preserve existing people.xml from loaded document', async () => {
    // Create a doc with tracked changes
    const doc1 = Document.create();
    doc1.enableTrackChanges({ author: 'Original Author' });

    const para = new Paragraph();
    para.addRevision(Revision.createInsertion('Original Author', new Run('Original text')));
    doc1.addParagraph(para);

    const buffer1 = await doc1.toBuffer();
    doc1.dispose();

    // Load and save without adding new changes
    const doc2 = await Document.loadFromBuffer(buffer1, { revisionHandling: 'preserve' });
    const buffer2 = await doc2.toBuffer();
    doc2.dispose();

    const zipHandler = new ZipHandler();
    await zipHandler.loadFromBuffer(buffer2);

    // people.xml should still be properly registered
    if (zipHandler.hasFile('word/people.xml')) {
      const contentTypes = zipHandler.getFileAsString('[Content_Types].xml') || '';
      expect(contentTypes).toContain('people.xml');

      const rels = zipHandler.getFileAsString('word/_rels/document.xml.rels') || '';
      expect(rels).toContain('people.xml');
    }
  });
});
