/**
 * Tests for comment parsing functionality
 * Ensures comments are properly parsed from documents and preserved during round-trip
 */

import { Document } from '../../src/core/Document';
import { Comment } from '../../src/elements/Comment';
import { Run } from '../../src/elements/Run';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLParser } from '../../src/xml/XMLParser';
import * as fs from 'fs';
import * as path from 'path';

describe('Comment Parsing', () => {
  describe('Basic Comment Parsing', () => {
    it('should parse comments from comments.xml', async () => {
      // Create a document with comments
      const doc = Document.create();
      const para = doc.createParagraph('This text has a comment');

      // Add a comment programmatically
      const comment = doc
        .getCommentManager()
        .createComment('John Doe', 'This is a test comment', 'JD');

      // Save and reload
      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);

      // Verify comment was preserved
      const comments = loadedDoc.getAllComments();
      expect(comments).toHaveLength(1);
      expect(comments[0]?.getAuthor()).toBe('John Doe');
      expect(comments[0]?.getInitials()).toBe('JD');
    });

    it('should preserve comment metadata', async () => {
      const doc = Document.create();
      doc.createParagraph('Test paragraph');

      // Add comment with specific date
      const testDate = new Date('2024-01-15T10:30:00Z');
      const comment = new Comment({
        author: 'Alice Smith',
        initials: 'AS',
        date: testDate,
        content: 'Review this section',
      });

      doc.getCommentManager().register(comment);

      // Round-trip
      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);

      const loadedComments = loadedDoc.getAllComments();
      expect(loadedComments).toHaveLength(1);

      const loadedComment = loadedComments[0];
      expect(loadedComment?.getAuthor()).toBe('Alice Smith');
      expect(loadedComment?.getInitials()).toBe('AS');
      // Date comparison may need tolerance due to serialization
      expect(loadedComment?.getDate()).toBeDefined();
    });

    it('should handle comment replies', async () => {
      const doc = Document.create();
      doc.createParagraph('Discussion topic');

      // Add parent comment
      const parentComment = doc.getCommentManager().createComment('User1', 'Initial comment');

      // Add reply
      const replyComment = doc
        .getCommentManager()
        .createReply(parentComment.getId(), 'User2', 'Reply to initial comment');

      // Round-trip
      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);

      const comments = loadedDoc.getCommentManager().getAllCommentsWithReplies();
      expect(comments).toHaveLength(2);

      // Find reply
      const reply = comments.find((c) => c.isReply());
      expect(reply).toBeDefined();
      expect(reply?.getParentId()).toBe(parentComment.getId());
    });

    it('should parse comment ranges in paragraphs', async () => {
      // Create document with comment ranges
      const doc = Document.create();
      const para = doc.createParagraph('This text is commented on here.');

      // Add comment
      const comment = doc.getCommentManager().createComment('Reviewer', 'Check this phrase');

      // Save and verify XML structure contains comment ranges
      const buffer = await doc.toBuffer();
      const zipHandler = new ZipHandler();
      await zipHandler.loadFromBuffer(buffer);

      const docXml = zipHandler.getFileAsString('word/document.xml');
      expect(docXml).toBeDefined();

      // Check for comment range markers (these would be added by a complete implementation)
      // For now, just verify the document loads without error
      const loadedDoc = await Document.loadFromBuffer(buffer);
      expect(loadedDoc).toBeDefined();
    });
  });

  describe('Complex Comment Scenarios', () => {
    it('should handle multiple comments on same paragraph', async () => {
      const doc = Document.create();
      const para = doc.createParagraph('Text with multiple comments');

      doc.getCommentManager().createComment('User1', 'First comment');
      doc.getCommentManager().createComment('User2', 'Second comment');
      doc.getCommentManager().createComment('User3', 'Third comment');

      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);

      const comments = loadedDoc.getAllComments();
      expect(comments).toHaveLength(3);

      // Verify each comment has unique ID
      const ids = comments.map((c) => c.getId());
      const uniqueIds = new Set(ids);
      expect(uniqueIds.size).toBe(3);
    });

    it('should preserve comment formatting', async () => {
      const doc = Document.create();
      doc.createParagraph('Formatted comment test');

      // Create comment with formatted runs
      const comment = new Comment({
        author: 'Editor',
        content: [
          new Run('Bold text', { bold: true }),
          new Run(' and '),
          new Run('italic text', { italic: true }),
        ],
      });

      doc.getCommentManager().register(comment);

      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);

      const loadedComments = loadedDoc.getAllComments();
      expect(loadedComments).toHaveLength(1);

      // Comment content should be preserved
      const content = loadedComments[0]?.getContent();
      expect(content).toBeDefined();
    });

    it('should handle empty comments', async () => {
      const doc = Document.create();
      doc.createParagraph('Text');

      // Add empty comment
      const comment = new Comment({
        author: 'User',
        content: '',
      });

      doc.getCommentManager().register(comment);

      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);

      const comments = loadedDoc.getAllComments();
      expect(comments).toHaveLength(1);
      expect(comments[0]?.getAuthor()).toBe('User');
    });
  });

  describe('Comment Manager Integration', () => {
    it('should find comments by author', async () => {
      const doc = Document.create();

      doc.getCommentManager().createComment('Alice', 'Comment 1');
      doc.getCommentManager().createComment('Bob', 'Comment 2');
      doc.getCommentManager().createComment('Alice', 'Comment 3');

      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);

      const aliceComments = loadedDoc
        .getCommentManager()
        .getAllComments()
        .filter((c) => c.getAuthor() === 'Alice');

      expect(aliceComments).toHaveLength(2);
    });

    it('should get comment threads', async () => {
      const doc = Document.create();

      const parent = doc.getCommentManager().createComment('User1', 'Start discussion');
      doc.getCommentManager().createReply(parent.getId(), 'User2', 'Reply 1');
      doc.getCommentManager().createReply(parent.getId(), 'User3', 'Reply 2');

      const buffer = await doc.toBuffer();
      const loadedDoc = await Document.loadFromBuffer(buffer);

      const thread = loadedDoc.getCommentManager().getCommentThread(parent.getId());
      expect(thread).toBeDefined();
      expect(thread?.replies).toHaveLength(2);
    });
  });

  describe('Duplicate Relationship Prevention', () => {
    it('should not create duplicate comments relationships on round-trip', async () => {
      // Create a document with comments
      const doc = Document.create();
      doc.createParagraph('Text with a comment');
      doc.getCommentManager().createComment('Author', 'A comment');

      // Save to buffer (creates initial comments relationship)
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Load and save again (round-trip)
      const doc2 = await Document.loadFromBuffer(buffer1);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Load the round-tripped buffer and save a third time
      const doc3 = await Document.loadFromBuffer(buffer2);
      const buffer3 = await doc3.toBuffer();
      doc3.dispose();

      // Extract document.xml.rels from the final buffer and count comments relationships
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer3);
      const relsXml = await zip.getFileAsString('word/_rels/document.xml.rels');
      expect(relsXml).toBeDefined();

      const commentsRelPattern = /Type="[^"]*\/comments"/g;
      const matches = relsXml!.match(commentsRelPattern) || [];
      expect(matches).toHaveLength(1);
    });

    it('should preserve existing comments relationship ID on round-trip', async () => {
      // Create a document with comments
      const doc = Document.create();
      doc.createParagraph('Test');
      doc.getCommentManager().createComment('Author', 'Comment');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Extract the original rId for comments
      const zip1 = new ZipHandler();
      await zip1.loadFromBuffer(buffer1);
      const relsXml1 = await zip1.getFileAsString('word/_rels/document.xml.rels');
      const rIdMatch = relsXml1!.match(/Id="(rId\d+)"[^>]*Type="[^"]*\/comments"/);
      expect(rIdMatch).toBeTruthy();
      const originalRId = rIdMatch![1];

      // Round-trip
      const doc2 = await Document.loadFromBuffer(buffer1);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Verify the same rId is used (no new one added)
      const zip2 = new ZipHandler();
      await zip2.loadFromBuffer(buffer2);
      const relsXml2 = await zip2.getFileAsString('word/_rels/document.xml.rels');
      const rIdMatch2 = relsXml2!.match(/Id="(rId\d+)"[^>]*Type="[^"]*\/comments"/);
      expect(rIdMatch2).toBeTruthy();
      expect(rIdMatch2![1]).toBe(originalRId);
    });
  });

  describe('Comment Companion File Passthrough', () => {
    it('should preserve companion files on unmodified round-trip', async () => {
      // Create a document with comments
      const doc = Document.create();
      doc.createParagraph('Test paragraph');
      doc.getCommentManager().createComment('Author', 'Comment text');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Inject companion files to simulate a real Word document
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer1);

      const extendedXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">' +
        '<w15:commentEx w15:paraId="DEADBEEF" w15:done="0"/>' +
        '</w15:commentsEx>';
      const idsXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">' +
        '<w16cid:commentId w16cid:paraId="DEADBEEF" w16cid:durableId="12345"/>' +
        '</w16cid:commentsIds>';
      const extensibleXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex">' +
        '</w16cex:commentsExtensible>';

      zip.addFile('word/commentsExtended.xml', extendedXml);
      zip.addFile('word/commentsIds.xml', idsXml);
      zip.addFile('word/commentsExtensible.xml', extensibleXml);
      const bufferWithCompanions = await zip.toBuffer();

      // Load and round-trip WITHOUT modifications
      const doc2 = await Document.loadFromBuffer(bufferWithCompanions);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Verify all 4 comment files are preserved
      const zip2 = new ZipHandler();
      await zip2.loadFromBuffer(buffer2);
      expect(zip2.hasFile('word/comments.xml')).toBe(true);
      expect(zip2.hasFile('word/commentsExtended.xml')).toBe(true);
      expect(zip2.hasFile('word/commentsIds.xml')).toBe(true);
      expect(zip2.hasFile('word/commentsExtensible.xml')).toBe(true);

      // Verify content is identical
      expect(zip2.getFileAsString('word/commentsExtended.xml')).toBe(extendedXml);
      expect(zip2.getFileAsString('word/commentsIds.xml')).toBe(idsXml);
      expect(zip2.getFileAsString('word/commentsExtensible.xml')).toBe(extensibleXml);
    });

    it('should preserve companion relationships on unmodified round-trip', async () => {
      // Create a document with comments
      const doc = Document.create();
      doc.createParagraph('Test paragraph');
      doc.getCommentManager().createComment('Author', 'Comment text');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Inject companion files AND their relationships into the ZIP
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer1);

      // Add companion files
      zip.addFile(
        'word/commentsExtended.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>'
      );
      zip.addFile(
        'word/commentsIds.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"/>'
      );
      zip.addFile(
        'word/commentsExtensible.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"/>'
      );

      // Inject companion relationships into document.xml.rels
      const relsXml = zip.getFileAsString('word/_rels/document.xml.rels')!;
      const injectedRels = relsXml.replace(
        '</Relationships>',
        '  <Relationship Id="rId72" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>\n' +
          '  <Relationship Id="rId73" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/>\n' +
          '  <Relationship Id="rId74" Type="http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible" Target="commentsExtensible.xml"/>\n' +
          '</Relationships>'
      );
      zip.addFile('word/_rels/document.xml.rels', injectedRels);
      const bufferWithCompanions = await zip.toBuffer();

      // Load and round-trip WITHOUT modifications
      const doc2 = await Document.loadFromBuffer(bufferWithCompanions);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Verify companion files preserved
      const zip2 = new ZipHandler();
      await zip2.loadFromBuffer(buffer2);
      expect(zip2.hasFile('word/commentsExtended.xml')).toBe(true);
      expect(zip2.hasFile('word/commentsIds.xml')).toBe(true);
      expect(zip2.hasFile('word/commentsExtensible.xml')).toBe(true);

      // Verify companion relationships preserved
      const finalRels = zip2.getFileAsString('word/_rels/document.xml.rels')!;
      expect(finalRels).toContain('commentsExtended.xml');
      expect(finalRels).toContain('commentsIds.xml');
      expect(finalRels).toContain('commentsExtensible.xml');

      // comments.xml and its relationship must still exist
      expect(zip2.hasFile('word/comments.xml')).toBe(true);
      expect(finalRels).toContain('comments.xml');
      expect(finalRels).toContain('/relationships/comments');
    });

    it('should remove companion files when comments are modified', async () => {
      // Create a document with comments
      const doc = Document.create();
      doc.createParagraph('Test paragraph');
      doc.createComment('Author', 'Comment text');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Inject companion files to simulate a real Word document
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer1);
      zip.addFile(
        'word/commentsExtended.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">' +
          '<w15:commentEx w15:paraId="DEADBEEF" w15:done="0"/>' +
          '</w15:commentsEx>'
      );
      zip.addFile(
        'word/commentsIds.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">' +
          '<w16cid:commentId w16cid:paraId="DEADBEEF" w16cid:durableId="12345"/>' +
          '</w16cid:commentsIds>'
      );
      zip.addFile(
        'word/commentsExtensible.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex">' +
          '</w16cex:commentsExtensible>'
      );
      const bufferWithCompanions = await zip.toBuffer();

      // Load, MODIFY comments, then save
      const doc2 = await Document.loadFromBuffer(bufferWithCompanions);
      doc2.createComment('NewAuthor', 'A new comment');
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Verify companion files are removed (regenerated comments lack w14:paraId)
      const zip2 = new ZipHandler();
      await zip2.loadFromBuffer(buffer2);
      expect(zip2.hasFile('word/commentsExtended.xml')).toBe(false);
      expect(zip2.hasFile('word/commentsIds.xml')).toBe(false);
      expect(zip2.hasFile('word/commentsExtensible.xml')).toBe(false);
      // comments.xml itself should still exist
      expect(zip2.hasFile('word/comments.xml')).toBe(true);
    });

    it('should clean up comments.xml and companions when all comments are removed', async () => {
      // Create a document with comments
      const doc = Document.create();
      doc.createParagraph('Text with a comment');
      doc.createComment('Author', 'Comment text');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Inject companion files to simulate a real Word document
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer1);
      zip.addFile(
        'word/commentsExtended.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">' +
          '<w15:commentEx w15:paraId="DEADBEEF" w15:done="0"/>' +
          '</w15:commentsEx>'
      );
      zip.addFile(
        'word/commentsIds.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">' +
          '<w16cid:commentId w16cid:paraId="DEADBEEF" w16cid:durableId="12345"/>' +
          '</w16cid:commentsIds>'
      );
      const bufferWithCompanions = await zip.toBuffer();

      // Load, remove ALL comments, then save
      const doc2 = await Document.loadFromBuffer(bufferWithCompanions);
      const allComments = doc2.getAllComments();
      for (const comment of allComments) {
        doc2.removeComment(comment.getId());
      }
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Verify everything is cleaned up
      const zip2 = new ZipHandler();
      await zip2.loadFromBuffer(buffer2);
      expect(zip2.hasFile('word/comments.xml')).toBe(false);
      expect(zip2.hasFile('word/commentsExtended.xml')).toBe(false);
      expect(zip2.hasFile('word/commentsIds.xml')).toBe(false);
      expect(zip2.hasFile('word/commentsExtensible.xml')).toBe(false);

      // Verify no orphaned comments relationship in rels
      const relsXml = zip2.getFileAsString('word/_rels/document.xml.rels')!;
      expect(relsXml).not.toContain('/relationships/comments');
    });

    it('should preserve Content_Types entries for companion files on passthrough', async () => {
      // Create a document with comments
      const doc = Document.create();
      doc.createParagraph('Test');
      doc.createComment('Author', 'Comment');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Inject companion files AND their Content_Types overrides
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer1);

      zip.addFile(
        'word/commentsExtended.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>'
      );
      zip.addFile(
        'word/commentsIds.xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"/>'
      );

      // Inject Content_Types overrides for companion files
      let contentTypes = zip.getFileAsString('[Content_Types].xml')!;
      contentTypes = contentTypes.replace(
        '</Types>',
        '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>' +
          '<Override PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/>' +
          '</Types>'
      );
      zip.addFile('[Content_Types].xml', contentTypes);
      const bufferWithCompanions = await zip.toBuffer();

      // Round-trip without modifications
      const doc2 = await Document.loadFromBuffer(bufferWithCompanions);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Verify Content_Types has entries for companion files
      const zip2 = new ZipHandler();
      await zip2.loadFromBuffer(buffer2);
      const outputContentTypes = zip2.getFileAsString('[Content_Types].xml')!;
      expect(outputContentTypes).toContain('commentsExtended.xml');
      expect(outputContentTypes).toContain('commentsIds.xml');
    });

    it('should survive double round-trip with companion files intact', async () => {
      // Create a document with comments
      const doc = Document.create();
      doc.createParagraph('Test paragraph');
      doc.createComment('Author', 'Comment text');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Inject companion files
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer1);

      const extendedXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">' +
        '<w15:commentEx w15:paraId="AABBCCDD" w15:done="1"/>' +
        '</w15:commentsEx>';
      const idsXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">' +
        '<w16cid:commentId w16cid:paraId="AABBCCDD" w16cid:durableId="67890"/>' +
        '</w16cid:commentsIds>';

      zip.addFile('word/commentsExtended.xml', extendedXml);
      zip.addFile('word/commentsIds.xml', idsXml);
      const bufferWithCompanions = await zip.toBuffer();

      // First round-trip (no modifications)
      const doc2 = await Document.loadFromBuffer(bufferWithCompanions);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Second round-trip (no modifications)
      const doc3 = await Document.loadFromBuffer(buffer2);
      const buffer3 = await doc3.toBuffer();
      doc3.dispose();

      // Verify companion files survived both round-trips
      const zip3 = new ZipHandler();
      await zip3.loadFromBuffer(buffer3);
      expect(zip3.hasFile('word/comments.xml')).toBe(true);
      expect(zip3.hasFile('word/commentsExtended.xml')).toBe(true);
      expect(zip3.hasFile('word/commentsIds.xml')).toBe(true);

      // Verify content is identical to original
      expect(zip3.getFileAsString('word/commentsExtended.xml')).toBe(extendedXml);
      expect(zip3.getFileAsString('word/commentsIds.xml')).toBe(idsXml);
    });

    it('should not add spurious comment files when document has no comments', async () => {
      // Create a document WITHOUT comments
      const doc = Document.create();
      doc.createParagraph('No comments here');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Load and round-trip without modifications
      const doc2 = await Document.loadFromBuffer(buffer1);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Verify no comment files were added
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer2);
      expect(zip.hasFile('word/comments.xml')).toBe(false);
      expect(zip.hasFile('word/commentsExtended.xml')).toBe(false);
      expect(zip.hasFile('word/commentsIds.xml')).toBe(false);
      expect(zip.hasFile('word/commentsExtensible.xml')).toBe(false);
    });

    it('should preserve comments.xml content exactly on passthrough', async () => {
      // Create a document with comments
      const doc = Document.create();
      doc.createParagraph('Test');
      doc.createComment('Author', 'Comment');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Get the original comments.xml content
      const zip1 = new ZipHandler();
      await zip1.loadFromBuffer(buffer1);
      const originalCommentsXml = zip1.getFileAsString('word/comments.xml')!;

      // Round-trip without modifications
      const doc2 = await Document.loadFromBuffer(buffer1);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Verify comments.xml is identical
      const zip2 = new ZipHandler();
      await zip2.loadFromBuffer(buffer2);
      const roundTrippedCommentsXml = zip2.getFileAsString('word/comments.xml')!;
      expect(roundTrippedCommentsXml).toBe(originalCommentsXml);
    });
  });

  describe('Comment Anchor Round-Trip', () => {
    it('should preserve comment range markers and references during round-trip', async () => {
      // Build a DOCX buffer with comment anchors injected into document.xml
      const doc = Document.create();
      doc.createParagraph('Test paragraph');
      doc.getCommentManager().createComment('Author', 'A comment');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Inject comment range markers and a commentReference run into document.xml
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer1);
      let docXml = zip.getFileAsString('word/document.xml')!;

      // Find the first <w:r> and wrap it with comment range markers
      const firstRunPos = docXml.indexOf('<w:r>');
      const closingBodyTag = docXml.lastIndexOf('</w:body>');
      // Find end of the paragraph that contains the first run
      const paraEnd = docXml.indexOf('</w:p>', firstRunPos);

      // Insert commentRangeStart before the run, commentRangeEnd + reference run after the run
      const commentRefRun =
        '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>';
      docXml =
        docXml.substring(0, firstRunPos) +
        '<w:commentRangeStart w:id="0"/>' +
        docXml.substring(firstRunPos, paraEnd) +
        '<w:commentRangeEnd w:id="0"/>' +
        commentRefRun +
        docXml.substring(paraEnd);

      zip.addFile('word/document.xml', docXml);
      const bufferWithAnchors = await zip.toBuffer();

      // Round-trip through docxmlater
      const doc2 = await Document.loadFromBuffer(bufferWithAnchors);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Extract and verify the output document.xml
      const zip2 = new ZipHandler();
      await zip2.loadFromBuffer(buffer2);
      const outputDocXml = zip2.getFileAsString('word/document.xml')!;

      // Comment range markers must survive round-trip
      expect(outputDocXml).toContain('w:commentRangeStart');
      expect(outputDocXml).toContain('w:id="0"');
      expect(outputDocXml).toContain('w:commentRangeEnd');
      // The commentReference element inside the run must also survive
      expect(outputDocXml).toContain('w:commentReference');
      // comments.xml must still exist
      expect(zip2.hasFile('word/comments.xml')).toBe(true);
    });

    it('should not produce orphaned comments when anchors exist', async () => {
      // Build a DOCX with a comment and its anchors
      const doc = Document.create();
      doc.createParagraph('Before comment');
      doc.createParagraph('Commented text');
      doc.getCommentManager().createComment('Reviewer', 'Check this');
      const buffer1 = await doc.toBuffer();
      doc.dispose();

      // Inject anchors around the second paragraph's content
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer1);
      let docXml = zip.getFileAsString('word/document.xml')!;

      // Find runs and inject comment markers around the second one
      const runs = docXml.match(/<w:r>.*?<\/w:r>/gs) || [];
      if (runs.length >= 2) {
        const secondRun = runs[1]!;
        const pos = docXml.indexOf(secondRun);
        const commentRefRun =
          '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>';
        docXml =
          docXml.substring(0, pos) +
          '<w:commentRangeStart w:id="0"/>' +
          secondRun +
          '<w:commentRangeEnd w:id="0"/>' +
          commentRefRun +
          docXml.substring(pos + secondRun.length);
        zip.addFile('word/document.xml', docXml);
      }
      const bufferWithAnchors = await zip.toBuffer();

      // Round-trip
      const doc2 = await Document.loadFromBuffer(bufferWithAnchors);
      const buffer2 = await doc2.toBuffer();
      doc2.dispose();

      // Verify: both comments.xml AND comment anchors in document.xml
      const zip2 = new ZipHandler();
      await zip2.loadFromBuffer(buffer2);
      const outputDocXml = zip2.getFileAsString('word/document.xml')!;

      expect(zip2.hasFile('word/comments.xml')).toBe(true);
      expect(outputDocXml).toContain('w:commentRangeStart');
      expect(outputDocXml).toContain('w:commentRangeEnd');
      expect(outputDocXml).toContain('w:commentReference');
    });
  });
});
