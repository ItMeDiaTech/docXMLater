/**
 * Comment regeneration fidelity (comments.xml).
 *
 * Whenever any comment is created/modified/removed, Document regenerates the
 * ENTIRE comments.xml from the in-memory model. That regeneration must not be
 * lossy for the untouched comments: run formatting (w:rPr) has to survive,
 * multi-paragraph comments must keep their <w:p> boundaries (collapsing them
 * concatenates words across the lost breaks), and the <w:annotationRef/>
 * reference-mark run must be preserved.
 */
import { Document } from '../../src/core/Document';
import { Run } from '../../src/elements/Run';
import { ZipHandler } from '../../src/zip/ZipHandler';

const MULTI_PARA_COMMENTS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:comment w:id="0" w:author="Alice" w:date="2024-01-15T10:30:00Z" w:initials="A">' +
  '<w:p>' +
  '<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>' +
  '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>' +
  '<w:r><w:rPr><w:b/></w:rPr><w:t>First para bold</w:t></w:r>' +
  '</w:p>' +
  '<w:p>' +
  '<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>' +
  '<w:r><w:t>Second para</w:t></w:r>' +
  '</w:p>' +
  '</w:comment>' +
  '</w:comments>';

/** Builds a DOCX whose comments.xml holds a two-paragraph formatted comment */
async function buildDocxWithMultiParagraphComment(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('Body text');
  seed.createComment('Seed', 'placeholder');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);
  zip.addFile('word/comments.xml', MULTI_PARA_COMMENTS_XML);
  return zip.toBuffer();
}

function extractCommentBlock(commentsXml: string, id: number): string {
  const match = commentsXml.match(new RegExp(`<w:comment w:id="${id}"[\\s\\S]*?</w:comment>`));
  expect(match).toBeTruthy();
  return match![0];
}

// CT_OnOff "on" forms: bare <w:b/> or explicit w:val="1"/"true"/"on"
const BOLD_ON = /<w:b(?: w:val="(?:1|true|on)")?\/>/;
const ITALIC_ON = /<w:i(?: w:val="(?:1|true|on)")?\/>/;

describe('X41: comments.xml regeneration fidelity', () => {
  it('keeps run formatting (w:rPr) when serializing programmatic comments', async () => {
    const doc = Document.create();
    try {
      doc.createParagraph('Body');
      doc.createComment(
        'Alice',
        new Run('This needs attention', { bold: true, color: 'FF0000' }),
        'A'
      );
      const buffer = await doc.toBuffer();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const commentsXml = zip.getFileAsString('word/comments.xml')!;

      expect(commentsXml).toContain('This needs attention');
      expect(commentsXml).toMatch(BOLD_ON);
      expect(commentsXml).toMatch(/<w:color w:val="FF0000"/);
    } finally {
      doc.dispose();
    }
  });

  it('preserves paragraph boundaries, formatting, and annotationRef of parsed comments across a modify-save', async () => {
    const crafted = await buildDocxWithMultiParagraphComment();
    const doc = await Document.loadFromBuffer(crafted);
    try {
      // Any comment mutation forces full regeneration of comments.xml
      doc.createComment('Bob', 'trigger regeneration');
      const out = await doc.toBuffer();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const commentsXml = zip.getFileAsString('word/comments.xml')!;
      const block = extractCommentBlock(commentsXml, 0);

      const paragraphs = block.match(/<w:p>[\s\S]*?<\/w:p>/g) || [];
      expect(paragraphs).toHaveLength(2);

      // First paragraph: reference mark run + bold run
      expect(paragraphs[0]).toContain('<w:annotationRef/>');
      expect(paragraphs[0]).toContain('<w:rStyle w:val="CommentReference"/>');
      expect(paragraphs[0]).toMatch(BOLD_ON);
      expect(paragraphs[0]).toContain('First para bold');
      expect(paragraphs[0]).not.toContain('Second para');

      // Second paragraph keeps its own text, not merged into the first
      expect(paragraphs[1]).toContain('Second para');
      expect(paragraphs[1]).not.toContain('First para bold');
    } finally {
      doc.dispose();
    }
  });

  it('serializes runs added after parsing into the last original paragraph', async () => {
    const crafted = await buildDocxWithMultiParagraphComment();
    const doc = await Document.loadFromBuffer(crafted);
    try {
      const comment = doc.getAllComments()[0]!;
      comment.addRun(new Run(' appended', { italic: true }));
      doc.createComment('Bob', 'trigger regeneration');
      const out = await doc.toBuffer();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const commentsXml = zip.getFileAsString('word/comments.xml')!;
      const block = extractCommentBlock(commentsXml, 0);

      const paragraphs = block.match(/<w:p>[\s\S]*?<\/w:p>/g) || [];
      expect(paragraphs).toHaveLength(2);
      expect(paragraphs[1]).toContain('Second para');
      expect(paragraphs[1]).toContain(' appended');
      expect(paragraphs[1]).toMatch(ITALIC_ON);
    } finally {
      doc.dispose();
    }
  });
});
