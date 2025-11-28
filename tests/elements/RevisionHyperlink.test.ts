/**
 * RevisionHyperlink Tests
 *
 * Tests for tracked changes support for hyperlinks including:
 * - Revision class with Hyperlink content
 * - Hyperlink clone() method
 * - Paragraph replaceContent() method
 * - XML generation for deleted/inserted hyperlinks
 * - Round-trip preservation
 */

import { describe, it, expect, beforeAll } from '@jest/globals';
import { join } from 'path';
import { promises as fs } from 'fs';
import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { Revision } from '../../src/elements/Revision';
import { Run } from '../../src/elements/Run';
import { isRunContent, isHyperlinkContent } from '../../src/elements/RevisionContent';

const OUTPUT_DIR = join(__dirname, '../output');

// Ensure output directory exists
beforeAll(async () => {
  try {
    await fs.mkdir(OUTPUT_DIR, { recursive: true });
  } catch {
    // Directory may already exist
  }
});

describe('RevisionHyperlink Tests', () => {
  describe('Hyperlink clone() method', () => {
    it('should clone a hyperlink with URL', () => {
      const original = new Hyperlink({
        url: 'https://example.com',
        text: 'Example Link',
        tooltip: 'Click here',
      });

      const cloned = original.clone();

      expect(cloned.getUrl()).toBe('https://example.com');
      expect(cloned.getText()).toBe('Example Link');
      expect(cloned.getTooltip()).toBe('Click here');
      expect(cloned).not.toBe(original); // Different instances
    });

    it('should clone a hyperlink with anchor', () => {
      const original = new Hyperlink({
        anchor: 'bookmark1',
        text: 'Internal Link',
      });

      const cloned = original.clone();

      expect(cloned.getAnchor()).toBe('bookmark1');
      expect(cloned.getText()).toBe('Internal Link');
    });

    it('should clone hyperlink formatting', () => {
      const original = new Hyperlink({
        url: 'https://example.com',
        text: 'Styled Link',
        formatting: { bold: true, color: 'FF0000' },
      });

      const cloned = original.clone();
      const formatting = cloned.getRawFormatting();

      expect(formatting?.bold).toBe(true);
      expect(formatting?.color).toBe('FF0000');
    });

    it('should create independent copy (modifying clone does not affect original)', () => {
      const original = new Hyperlink({
        url: 'https://old-url.com',
        text: 'Link',
      });

      const cloned = original.clone();
      cloned.setUrl('https://new-url.com');

      expect(original.getUrl()).toBe('https://old-url.com');
      expect(cloned.getUrl()).toBe('https://new-url.com');
    });
  });

  describe('Revision with Hyperlink content', () => {
    it('should create insertion revision with hyperlink', () => {
      const hyperlink = new Hyperlink({
        url: 'https://example.com',
        text: 'New Link',
      });

      const revision = Revision.createInsertion('TestAuthor', [hyperlink]);

      expect(revision.getType()).toBe('insert');
      expect(revision.getAuthor()).toBe('TestAuthor');

      const hyperlinks = revision.getHyperlinks();
      expect(hyperlinks).toHaveLength(1);
      expect(hyperlinks[0]!.getText()).toBe('New Link');
    });

    it('should create deletion revision with hyperlink', () => {
      const hyperlink = new Hyperlink({
        url: 'https://old-url.com',
        text: 'Deleted Link',
      });

      const revision = Revision.createDeletion('TestAuthor', [hyperlink]);

      expect(revision.getType()).toBe('delete');
      const hyperlinks = revision.getHyperlinks();
      expect(hyperlinks).toHaveLength(1);
    });

    it('should add hyperlink to revision', () => {
      const revision = Revision.createInsertion('TestAuthor', []);

      const hyperlink = new Hyperlink({
        url: 'https://added.com',
        text: 'Added Link',
      });

      revision.addHyperlink(hyperlink);

      expect(revision.getHyperlinks()).toHaveLength(1);
    });

    it('should get content including both runs and hyperlinks', () => {
      const run = new Run('Some text');
      const hyperlink = new Hyperlink({
        url: 'https://example.com',
        text: 'Link',
      });

      const revision = Revision.createInsertion('TestAuthor', [run, hyperlink]);
      const content = revision.getContent();

      expect(content).toHaveLength(2);
      expect(isRunContent(content[0]!)).toBe(true);
      expect(isHyperlinkContent(content[1]!)).toBe(true);
    });
  });

  describe('Type guards for RevisionContent', () => {
    it('should identify Run content', () => {
      const run = new Run('Text');
      expect(isRunContent(run)).toBe(true);
      expect(isHyperlinkContent(run)).toBe(false);
    });

    it('should identify Hyperlink content', () => {
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      expect(isHyperlinkContent(hyperlink)).toBe(true);
      expect(isRunContent(hyperlink)).toBe(false);
    });
  });

  describe('Paragraph replaceContent() method', () => {
    it('should replace single item with multiple items', () => {
      const para = new Paragraph();
      const hyperlink = new Hyperlink({
        url: 'https://old.com',
        text: 'Old Link',
      });
      para.addHyperlink(hyperlink);

      const deletion = Revision.createDeletion('Author', [hyperlink.clone()]);
      const newHyperlink = new Hyperlink({
        url: 'https://new.com',
        text: 'New Link',
      });
      const insertion = Revision.createInsertion('Author', [newHyperlink]);

      const replaced = para.replaceContent(hyperlink, [deletion, insertion]);

      expect(replaced).toBe(true);
      const content = para.getContent();
      expect(content).toHaveLength(2);
      expect(content[0]).toBeInstanceOf(Revision);
      expect(content[1]).toBeInstanceOf(Revision);
    });

    it('should return false when item not found', () => {
      const para = new Paragraph();
      para.addText('Some text');

      const notInPara = new Hyperlink({ url: 'https://example.com', text: 'Not in para' });
      const replaced = para.replaceContent(notInPara, []);

      expect(replaced).toBe(false);
    });

    it('should maintain correct order after replacement', () => {
      const para = new Paragraph();
      para.addText('Before ');
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      para.addHyperlink(hyperlink);
      para.addText(' After');

      const replacement = new Run('Replaced');
      para.replaceContent(hyperlink, [replacement]);

      const content = para.getContent();
      expect(content).toHaveLength(3);
      expect((content[0] as Run).getText()).toBe('Before ');
      expect((content[1] as Run).getText()).toBe('Replaced');
      expect((content[2] as Run).getText()).toBe(' After');
    });
  });

  describe('Paragraph setContent() method', () => {
    it('should replace all content', () => {
      const para = new Paragraph();
      para.addText('Original text');
      para.addText('More text');

      const newRun = new Run('Completely new content');
      para.setContent([newRun]);

      const content = para.getContent();
      expect(content).toHaveLength(1);
      expect((content[0] as Run).getText()).toBe('Completely new content');
    });
  });

  describe('XML generation for revision hyperlinks', () => {
    it('should generate XML for insertion revision with hyperlink', () => {
      const hyperlink = new Hyperlink({
        url: 'https://example.com',
        text: 'New Link',
        relationshipId: 'rId1',
      });

      const revision = Revision.createInsertion('TestAuthor', [hyperlink]);
      const xml = revision.toXML();

      expect(xml.name).toBe('w:ins');
      expect(xml.attributes?.['w:author']).toBe('TestAuthor');
      expect(xml.children).toBeDefined();
      // Should contain the hyperlink
      const hyperlinkChild = xml.children?.find(
        (c) => typeof c === 'object' && c.name === 'w:hyperlink'
      );
      expect(hyperlinkChild).toBeDefined();
    });

    it('should generate XML for deletion revision with hyperlink (w:delText)', () => {
      const hyperlink = new Hyperlink({
        url: 'https://old.com',
        text: 'Deleted Link',
        relationshipId: 'rId1',
      });

      const revision = Revision.createDeletion('TestAuthor', [hyperlink]);
      const xml = revision.toXML();

      expect(xml.name).toBe('w:del');
      // The hyperlink should be transformed to use w:delText
      const hyperlinkChild = xml.children?.find(
        (c) => typeof c === 'object' && c.name === 'w:hyperlink'
      ) as any;
      expect(hyperlinkChild).toBeDefined();

      // Check that runs inside use w:delText instead of w:t
      if (hyperlinkChild?.children) {
        const runChild = hyperlinkChild.children.find(
          (c: any) => typeof c === 'object' && c.name === 'w:r'
        );
        if (runChild?.children) {
          const delTextChild = runChild.children.find(
            (c: any) => typeof c === 'object' && c.name === 'w:delText'
          );
          expect(delTextChild).toBeDefined();
        }
      }
    });
  });

  describe('Document round-trip with revision hyperlinks', () => {
    it('should create and save document with tracked hyperlink changes', async () => {
      const doc = Document.create();
      doc.enableTrackChanges({ author: 'TestAuthor' });

      const para = new Paragraph();
      para.addText('Click here: ');

      // Create old hyperlink (deletion)
      const oldHyperlink = new Hyperlink({
        url: 'https://old-url.com',
        text: 'Old Link',
      });

      // Create new hyperlink (insertion)
      const newHyperlink = new Hyperlink({
        url: 'https://new-url.com',
        text: 'New Link',
      });

      // Create revisions
      const deletion = Revision.createDeletion('TestAuthor', [oldHyperlink]);
      const insertion = Revision.createInsertion('TestAuthor', [newHyperlink]);

      // Add revisions to paragraph
      para.addRevision(deletion);
      para.addRevision(insertion);

      doc.addParagraph(para);

      // Save to buffer
      const buffer = await doc.toBuffer();
      await fs.writeFile(join(OUTPUT_DIR, 'test-revision-hyperlink.docx'), buffer);

      expect(buffer).toBeDefined();
      expect(buffer.length).toBeGreaterThan(0);

      doc.dispose();
    });

    it('should handle hyperlink URL update with track changes', async () => {
      const doc = Document.create();
      doc.enableTrackChanges({ author: 'UpdateAuthor' });

      const para = new Paragraph();
      const hyperlink = new Hyperlink({
        url: 'https://original.com',
        text: 'Click Me',
      });
      para.addHyperlink(hyperlink);
      doc.addParagraph(para);

      // Simulate tracked URL change
      const oldHyperlink = hyperlink.clone();
      hyperlink.setUrl('https://updated.com');

      const deletion = Revision.createDeletion('UpdateAuthor', [oldHyperlink]);
      const insertion = Revision.createInsertion('UpdateAuthor', [hyperlink]);

      para.replaceContent(hyperlink, [deletion, insertion]);

      const buffer = await doc.toBuffer();
      await fs.writeFile(join(OUTPUT_DIR, 'test-hyperlink-url-update.docx'), buffer);

      expect(buffer).toBeDefined();
      doc.dispose();
    });
  });

  describe('Edge cases', () => {
    it('should handle empty revision content', () => {
      const revision = Revision.createInsertion('Author', []);
      expect(revision.getContent()).toHaveLength(0);
      expect(revision.getHyperlinks()).toHaveLength(0);
      expect(revision.getRuns()).toHaveLength(0);
    });

    it('should handle revision with mixed content types', () => {
      const run1 = new Run('Text before ');
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      const run2 = new Run(' text after');

      const revision = Revision.createInsertion('Author', [run1, hyperlink, run2]);

      expect(revision.getContent()).toHaveLength(3);
      expect(revision.getRuns()).toHaveLength(2);
      expect(revision.getHyperlinks()).toHaveLength(1);
    });

    it('should preserve hyperlink tooltip through clone', () => {
      const original = new Hyperlink({
        url: 'https://example.com',
        text: 'Hover me',
        tooltip: 'This is a tooltip',
      });

      const cloned = original.clone();
      expect(cloned.getTooltip()).toBe('This is a tooltip');
    });
  });
});
