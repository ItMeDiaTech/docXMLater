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

      expect(xml).not.toBeNull();
      // Per ECMA-376, w:hyperlink wraps w:ins (not the other way around)
      expect(xml!.name).toBe('w:hyperlink');
      expect(xml!.children).toBeDefined();
      // The w:ins should be inside the hyperlink
      const insChild = xml!.children?.find((c) => typeof c === 'object' && c.name === 'w:ins');
      expect(insChild).toBeDefined();
      if (insChild && typeof insChild !== 'string') {
        expect(insChild.attributes?.['w:author']).toBe('TestAuthor');
      }
    });

    it('should generate XML for deletion revision with hyperlink (w:delText)', () => {
      const hyperlink = new Hyperlink({
        url: 'https://old.com',
        text: 'Deleted Link',
        relationshipId: 'rId1',
      });

      const revision = Revision.createDeletion('TestAuthor', [hyperlink]);
      const xml = revision.toXML();

      expect(xml).not.toBeNull();
      // Per ECMA-376, w:hyperlink wraps w:del (not the other way around)
      expect(xml!.name).toBe('w:hyperlink');
      const delChild = xml!.children?.find(
        (c) => typeof c === 'object' && c.name === 'w:del'
      ) as any;
      expect(delChild).toBeDefined();

      // Check that runs inside use w:delText instead of w:t
      if (delChild?.children) {
        const runChild = delChild.children.find(
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

  describe('Automatic URL/Anchor tracking', () => {
    it('should create delete/insert pair when setUrl() called with tracking enabled', () => {
      // Create a mock tracking context
      const mockTrackingContext = {
        isEnabled: () => true,
        getAuthor: () => 'TestAuthor',
        trackHyperlinkChange: jest.fn(),
      };

      const para = new Paragraph();
      const hyperlink = new Hyperlink({ url: 'https://old.com', text: 'Link' });
      para.addHyperlink(hyperlink);

      // Enable tracking
      hyperlink._setTrackingContext(mockTrackingContext as any);

      // Verify parent paragraph is set
      expect(hyperlink._getParentParagraph()).toBe(para);

      // Change URL
      hyperlink.setUrl('https://new.com');

      // Verify paragraph now has revision pair
      const content = para.getContent();
      expect(content).toHaveLength(2);

      // First should be deletion
      expect(content[0]).toBeInstanceOf(Revision);
      const deletion = content[0] as Revision;
      expect(deletion.getType()).toBe('delete');
      const deletedHyperlinks = deletion.getHyperlinks();
      expect(deletedHyperlinks).toHaveLength(1);
      expect(deletedHyperlinks[0]!.getUrl()).toBe('https://old.com');

      // Second should be insertion
      expect(content[1]).toBeInstanceOf(Revision);
      const insertion = content[1] as Revision;
      expect(insertion.getType()).toBe('insert');
      const insertedHyperlinks = insertion.getHyperlinks();
      expect(insertedHyperlinks).toHaveLength(1);
      expect(insertedHyperlinks[0]!.getUrl()).toBe('https://new.com');
    });

    it('should create delete/insert pair when setAnchor() called with tracking enabled', () => {
      const mockTrackingContext = {
        isEnabled: () => true,
        getAuthor: () => 'TestAuthor',
        trackHyperlinkChange: jest.fn(),
      };

      const para = new Paragraph();
      const hyperlink = new Hyperlink({ anchor: 'oldBookmark', text: 'Link' });
      para.addHyperlink(hyperlink);

      hyperlink._setTrackingContext(mockTrackingContext as any);
      hyperlink.setAnchor('newBookmark');

      const content = para.getContent();
      expect(content).toHaveLength(2);

      const deletion = content[0] as Revision;
      expect(deletion.getType()).toBe('delete');
      const deletedHyperlinks = deletion.getHyperlinks();
      expect(deletedHyperlinks).toHaveLength(1);
      expect(deletedHyperlinks[0]!.getAnchor()).toBe('oldBookmark');

      const insertion = content[1] as Revision;
      expect(insertion.getType()).toBe('insert');
      const insertedHyperlinks = insertion.getHyperlinks();
      expect(insertedHyperlinks).toHaveLength(1);
      expect(insertedHyperlinks[0]!.getAnchor()).toBe('newBookmark');
    });

    it('should not track when tracking is disabled', () => {
      const mockTrackingContext = {
        isEnabled: () => false,
        getAuthor: () => 'TestAuthor',
        trackHyperlinkChange: jest.fn(),
      };

      const para = new Paragraph();
      const hyperlink = new Hyperlink({ url: 'https://old.com', text: 'Link' });
      para.addHyperlink(hyperlink);

      hyperlink._setTrackingContext(mockTrackingContext as any);
      hyperlink.setUrl('https://new.com');

      // Should still be just the hyperlink, no revisions
      const content = para.getContent();
      expect(content).toHaveLength(1);
      expect(content[0]).toBeInstanceOf(Hyperlink);
      expect((content[0] as Hyperlink).getUrl()).toBe('https://new.com');
    });

    it('should not track when no parent paragraph', () => {
      const mockTrackingContext = {
        isEnabled: () => true,
        getAuthor: () => 'TestAuthor',
        trackHyperlinkChange: jest.fn(),
      };

      const hyperlink = new Hyperlink({ url: 'https://old.com', text: 'Link' });
      // Note: NOT added to paragraph, so no parent reference

      hyperlink._setTrackingContext(mockTrackingContext as any);
      hyperlink.setUrl('https://new.com');

      // Should just update the URL without tracking (no parent to replace in)
      expect(hyperlink.getUrl()).toBe('https://new.com');
    });

    it('should clear parent reference after tracking', () => {
      const mockTrackingContext = {
        isEnabled: () => true,
        getAuthor: () => 'TestAuthor',
        trackHyperlinkChange: jest.fn(),
      };

      const para = new Paragraph();
      const hyperlink = new Hyperlink({ url: 'https://old.com', text: 'Link' });
      para.addHyperlink(hyperlink);
      hyperlink._setTrackingContext(mockTrackingContext as any);

      // After setUrl(), the hyperlink is inside a revision, no longer has parent
      hyperlink.setUrl('https://new.com');
      expect(hyperlink._getParentParagraph()).toBeUndefined();
    });

    it('should set parent reference when added via addHyperlink()', () => {
      const para = new Paragraph();

      // Test string overload
      const hyperlink1 = para.addHyperlink('https://example.com');
      expect(hyperlink1._getParentParagraph()).toBe(para);

      // Test Hyperlink object overload
      const hyperlink2 = new Hyperlink({ url: 'https://other.com', text: 'Other' });
      para.addHyperlink(hyperlink2);
      expect(hyperlink2._getParentParagraph()).toBe(para);

      // Test empty overload
      const hyperlink3 = para.addHyperlink();
      expect(hyperlink3._getParentParagraph()).toBe(para);
    });

    it('should set parent reference when using setContent()', () => {
      const para = new Paragraph();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      const run = new Run('Some text');

      para.setContent([run, hyperlink]);

      expect(hyperlink._getParentParagraph()).toBe(para);
    });
  });

  describe('Formatting tracked changes (rPrChange)', () => {
    function createTrackingMock() {
      let nextId = 100;
      return {
        isEnabled: () => true,
        getAuthor: () => 'FormatAuthor',
        getRevisionManager: () => ({
          consumeNextId: () => nextId++,
        }),
        trackHyperlinkChange: jest.fn(),
      };
    }

    it('should create rPrChange when setColor() called with tracking', () => {
      const mock = createTrackingMock();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      hyperlink._setTrackingContext(mock as any);

      // Default color is '0000FF'
      hyperlink.setColor('FF0000');

      const rPrChange = hyperlink.getRun().getPropertyChangeRevision();
      expect(rPrChange).toBeDefined();
      expect(rPrChange!.author).toBe('FormatAuthor');
      expect(rPrChange!.previousProperties.color).toBe('0000FF');
    });

    it('should create rPrChange when setBold() called with tracking', () => {
      const mock = createTrackingMock();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setBold(true);

      const rPrChange = hyperlink.getRun().getPropertyChangeRevision();
      expect(rPrChange).toBeDefined();
      expect(rPrChange!.previousProperties.bold).toBeUndefined();
    });

    it('should create rPrChange when setFormatting() called with tracking', () => {
      const mock = createTrackingMock();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setFormatting({ bold: true, italic: true, color: 'FF0000' });

      const rPrChange = hyperlink.getRun().getPropertyChangeRevision();
      expect(rPrChange).toBeDefined();
      expect(rPrChange!.previousProperties.color).toBe('0000FF');
    });

    it('should produce single merged rPrChange for multiple sequential formatting changes', () => {
      const mock = createTrackingMock();
      const hyperlink = new Hyperlink({
        url: 'https://example.com',
        text: 'Link',
        formatting: { color: '0000FF', bold: false },
      });
      hyperlink._setTrackingContext(mock as any);

      // Two sequential changes
      hyperlink.setColor('FF0000');
      hyperlink.setBold(true);

      const rPrChange = hyperlink.getRun().getPropertyChangeRevision();
      expect(rPrChange).toBeDefined();
      // Both previous values should be from the original baseline
      expect(rPrChange!.previousProperties.color).toBe('0000FF');
      expect(rPrChange!.previousProperties.bold).toBe(false);
    });

    it('should not create rPrChange when tracking is disabled', () => {
      const mock = {
        isEnabled: () => false,
        getAuthor: () => 'Author',
        getRevisionManager: () => ({ consumeNextId: () => 1 }),
        trackHyperlinkChange: jest.fn(),
      };
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setColor('FF0000');

      const rPrChange = hyperlink.getRun().getPropertyChangeRevision();
      expect(rPrChange).toBeUndefined();
    });

    it('should produce rPrChange in XML output with w:rPr > w:rPrChange structure', () => {
      const mock = createTrackingMock();
      const hyperlink = new Hyperlink({
        url: 'https://example.com',
        text: 'Link',
        relationshipId: 'rId1',
      });
      hyperlink._setTrackingContext(mock as any);

      // Change color from default '0000FF' to 'FF0000' — previous value is defined,
      // ensuring generateRunPropertiesXML produces non-null output for rPrChange
      hyperlink.setColor('FF0000');

      const xml = hyperlink.toXML();
      // The hyperlink should contain a run
      const runChild = xml.children?.find((c) => typeof c === 'object' && c.name === 'w:r') as any;
      expect(runChild).toBeDefined();

      // The run should have w:rPr with w:rPrChange inside
      const rPr = runChild?.children?.find((c: any) => typeof c === 'object' && c.name === 'w:rPr');
      expect(rPr).toBeDefined();
      if (rPr && typeof rPr === 'object') {
        const rPrChange = (rPr as any).children?.find(
          (c: any) => typeof c === 'object' && c.name === 'w:rPrChange'
        );
        expect(rPrChange).toBeDefined();
      }
    });

    it('should create rPrChange for setItalic()', () => {
      const mock = createTrackingMock();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setItalic(true);

      const rPrChange = hyperlink.getRun().getPropertyChangeRevision();
      expect(rPrChange).toBeDefined();
    });

    it('should create rPrChange for setFont()', () => {
      const mock = createTrackingMock();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setFont('Arial');

      const rPrChange = hyperlink.getRun().getPropertyChangeRevision();
      expect(rPrChange).toBeDefined();
      expect(rPrChange!.previousProperties.font).toBe('Verdana');
    });

    it('should create rPrChange for setSize()', () => {
      const mock = createTrackingMock();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setSize(24);

      const rPrChange = hyperlink.getRun().getPropertyChangeRevision();
      expect(rPrChange).toBeDefined();
      expect(rPrChange!.previousProperties.size).toBe(12);
    });

    it('should create rPrChange for setUnderline()', () => {
      const mock = createTrackingMock();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setUnderline('double');

      const rPrChange = hyperlink.getRun().getPropertyChangeRevision();
      expect(rPrChange).toBeDefined();
      expect(rPrChange!.previousProperties.underline).toBe('single');
    });
  });

  describe('Text tracked changes (delete/insert)', () => {
    function createTrackingMock() {
      return {
        isEnabled: () => true,
        getAuthor: () => 'TextAuthor',
        getRevisionManager: () => ({ consumeNextId: () => 200 }),
        trackHyperlinkChange: jest.fn(),
      };
    }

    it('should create delete/insert pair when setText() called with tracking', () => {
      const mock = createTrackingMock();
      const para = new Paragraph();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Old Text' });
      para.addHyperlink(hyperlink);
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setText('New Text');

      const content = para.getContent();
      expect(content).toHaveLength(2);

      const deletion = content[0] as Revision;
      expect(deletion.getType()).toBe('delete');
      expect(deletion.getHyperlinks()[0]!.getText()).toBe('Old Text');

      const insertion = content[1] as Revision;
      expect(insertion.getType()).toBe('insert');
      expect(insertion.getHyperlinks()[0]!.getText()).toBe('New Text');
    });

    it('should not create revision when text is unchanged', () => {
      const mock = createTrackingMock();
      const para = new Paragraph();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Same Text' });
      para.addHyperlink(hyperlink);
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setText('Same Text');

      // Should still be just the hyperlink, no revisions
      const content = para.getContent();
      expect(content).toHaveLength(1);
      expect(content[0]).toBeInstanceOf(Hyperlink);
    });

    it('should fall through to non-tracking when no parent paragraph', () => {
      const mock = createTrackingMock();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Old Text' });
      // NOT added to paragraph
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setText('New Text');

      // Should just update text without revisions
      expect(hyperlink.getText()).toBe('New Text');
    });

    it('should clear parent reference after setText() tracking', () => {
      const mock = createTrackingMock();
      const para = new Paragraph();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Old' });
      para.addHyperlink(hyperlink);
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setText('New');

      expect(hyperlink._getParentParagraph()).toBeUndefined();
    });
  });

  describe('Tooltip tracked changes (delete/insert)', () => {
    function createTrackingMock() {
      return {
        isEnabled: () => true,
        getAuthor: () => 'TooltipAuthor',
        getRevisionManager: () => ({ consumeNextId: () => 300 }),
        trackHyperlinkChange: jest.fn(),
      };
    }

    it('should create delete/insert pair when setTooltip() called with tracking', () => {
      const mock = createTrackingMock();
      const para = new Paragraph();
      const hyperlink = new Hyperlink({
        url: 'https://example.com',
        text: 'Link',
        tooltip: 'Old Tip',
      });
      para.addHyperlink(hyperlink);
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setTooltip('New Tip');

      const content = para.getContent();
      expect(content).toHaveLength(2);

      const deletion = content[0] as Revision;
      expect(deletion.getType()).toBe('delete');
      expect(deletion.getHyperlinks()[0]!.getTooltip()).toBe('Old Tip');

      const insertion = content[1] as Revision;
      expect(insertion.getType()).toBe('insert');
      expect(insertion.getHyperlinks()[0]!.getTooltip()).toBe('New Tip');
    });

    it('should not create revision when tooltip is unchanged', () => {
      const mock = createTrackingMock();
      const para = new Paragraph();
      const hyperlink = new Hyperlink({
        url: 'https://example.com',
        text: 'Link',
        tooltip: 'Same Tip',
      });
      para.addHyperlink(hyperlink);
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setTooltip('Same Tip');

      const content = para.getContent();
      expect(content).toHaveLength(1);
      expect(content[0]).toBeInstanceOf(Hyperlink);
    });

    it('should fall through to non-tracking when no parent paragraph', () => {
      const mock = createTrackingMock();
      const hyperlink = new Hyperlink({
        url: 'https://example.com',
        text: 'Link',
        tooltip: 'Old Tip',
      });
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setTooltip('New Tip');

      expect(hyperlink.getTooltip()).toBe('New Tip');
    });

    it('should clear parent reference after setTooltip() tracking', () => {
      const mock = createTrackingMock();
      const para = new Paragraph();
      const hyperlink = new Hyperlink({
        url: 'https://example.com',
        text: 'Link',
        tooltip: 'Old',
      });
      para.addHyperlink(hyperlink);
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setTooltip('New');

      expect(hyperlink._getParentParagraph()).toBeUndefined();
    });
  });

  describe('Integration: formatting + text changes', () => {
    it('should preserve rPrChange when followed by setText() with tracking', () => {
      let nextId = 400;
      const mock = {
        isEnabled: () => true,
        getAuthor: () => 'IntegrationAuthor',
        getRevisionManager: () => ({ consumeNextId: () => nextId++ }),
        trackHyperlinkChange: jest.fn(),
      };

      const para = new Paragraph();
      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      para.addHyperlink(hyperlink);
      hyperlink._setTrackingContext(mock as any);

      // First: formatting change (rPrChange on inner Run)
      hyperlink.setBold(true);
      expect(hyperlink.getRun().getPropertyChangeRevision()).toBeDefined();

      // Second: text change (delete/insert pair — clone preserves rPrChange)
      hyperlink.setText('New Text');

      const content = para.getContent();
      expect(content).toHaveLength(2);

      // Deletion should have the old text with rPrChange (cloned from modified hyperlink)
      const deletion = content[0] as Revision;
      const deletedLink = deletion.getHyperlinks()[0]!;
      expect(deletedLink.getText()).toBe('Link');
      expect(deletedLink.getRun().getPropertyChangeRevision()).toBeDefined();

      // Insertion should have the new text
      const insertion = content[1] as Revision;
      expect(insertion.getHyperlinks()[0]!.getText()).toBe('New Text');
    });

    it('should save valid buffer after tracked hyperlink formatting changes', async () => {
      const doc = Document.create();
      doc.enableTrackChanges({ author: 'SaveAuthor' });

      const para = new Paragraph();
      const hyperlink = new Hyperlink({
        url: 'https://example.com',
        text: 'Styled Link',
      });
      para.addHyperlink(hyperlink);
      doc.addParagraph(para);

      // Formatting change — produces rPrChange
      hyperlink.setBold(true);
      hyperlink.setColor('FF0000');

      const buffer = await doc.toBuffer();
      expect(buffer).toBeDefined();
      expect(buffer.length).toBeGreaterThan(0);

      doc.dispose();
    });

    it('should preserve rPrChange through clone()', () => {
      let nextId = 500;
      const mock = {
        isEnabled: () => true,
        getAuthor: () => 'CloneAuthor',
        getRevisionManager: () => ({ consumeNextId: () => nextId++ }),
        trackHyperlinkChange: jest.fn(),
      };

      const hyperlink = new Hyperlink({ url: 'https://example.com', text: 'Link' });
      hyperlink._setTrackingContext(mock as any);

      hyperlink.setBold(true);
      expect(hyperlink.getRun().getPropertyChangeRevision()).toBeDefined();

      const cloned = hyperlink.clone();
      const clonedRPrChange = cloned.getRun().getPropertyChangeRevision();
      expect(clonedRPrChange).toBeDefined();
      expect(clonedRPrChange!.author).toBe('CloneAuthor');
    });
  });
});
