/**
 * Comment regeneration must preserve per-paragraph properties and hyperlink
 * wrappers (comments.xml).
 *
 * CT_Comment content is EG_BlockLevelElts (ECMA-376 §17.13.4.2): Word emits
 * multi-paragraph comments where each paragraph carries the CommentText
 * pStyle, and links inside comments are w:hyperlink wrappers whose r:id
 * targets live in word/_rels/comments.xml.rels. When any comment is
 * added/removed, comments.xml is regenerated from the in-memory model, so
 * the parsed model must retain each paragraph's w:pPr and the hyperlink
 * wrappers — not just the flat run text.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

const COMMENTS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<w:comment w:id="0" w:author="Alice" w:date="2024-01-15T10:30:00Z" w:initials="A">' +
  '<w:p>' +
  '<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>' +
  '<w:r><w:t xml:space="preserve">See </w:t></w:r>' +
  '<w:hyperlink r:id="rId1" w:history="1">' +
  '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>the spec</w:t></w:r>' +
  '</w:hyperlink>' +
  '<w:r><w:t xml:space="preserve"> first</w:t></w:r>' +
  '</w:p>' +
  '<w:p>' +
  '<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>' +
  '<w:r><w:t>Second para</w:t></w:r>' +
  '</w:p>' +
  '</w:comment>' +
  '</w:comments>';

const COMMENTS_RELS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/spec" TargetMode="External"/>' +
  '</Relationships>';

/** Builds a DOCX whose comments.xml has a styled, hyperlinked, two-paragraph comment */
async function buildCraftedDocx(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('Body text');
  seed.createComment('Seed', 'placeholder');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);
  zip.addFile('word/comments.xml', COMMENTS_XML);
  zip.addFile('word/_rels/comments.xml.rels', COMMENTS_RELS_XML);
  return zip.toBuffer();
}

function extractCommentBlock(commentsXml: string, id: number): string {
  const match = commentsXml.match(new RegExp(`<w:comment w:id="${id}"[\\s\\S]*?</w:comment>`));
  expect(match).toBeTruthy();
  return match![0];
}

describe('C41: comment paragraph properties and hyperlink wrappers survive regeneration', () => {
  it('re-emits each paragraph w:pPr (CommentText pStyle) after a comment mutation forces regeneration', async () => {
    const crafted = await buildCraftedDocx();
    const doc = await Document.loadFromBuffer(crafted);
    try {
      doc.createComment('Bob', 'trigger regeneration');
      const out = await doc.toBuffer();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const commentsXml = zip.getFileAsString('word/comments.xml')!;
      const block = extractCommentBlock(commentsXml, 0);

      const paragraphs = block.match(/<w:p>[\s\S]*?<\/w:p>/g) || [];
      expect(paragraphs).toHaveLength(2);
      expect(paragraphs[0]).toContain('<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>');
      expect(paragraphs[1]).toContain('<w:pPr><w:pStyle w:val="CommentText"/></w:pPr>');
    } finally {
      doc.dispose();
    }
  });

  it('preserves w:hyperlink wrappers (r:id) and inline order within the paragraph', async () => {
    const crafted = await buildCraftedDocx();
    const doc = await Document.loadFromBuffer(crafted);
    try {
      doc.createComment('Bob', 'trigger regeneration');
      const out = await doc.toBuffer();

      const zip = new ZipHandler();
      await zip.loadFromBuffer(out);
      const commentsXml = zip.getFileAsString('word/comments.xml')!;
      const block = extractCommentBlock(commentsXml, 0);

      const paragraphs = block.match(/<w:p>[\s\S]*?<\/w:p>/g) || [];
      expect(paragraphs).toHaveLength(2);
      const first = paragraphs[0]!;

      // The hyperlink wrapper must survive verbatim, not be unwrapped to runs
      expect(first).toContain('<w:hyperlink r:id="rId1" w:history="1">');
      expect(first).toContain('</w:hyperlink>');
      const hyperlink = first.match(/<w:hyperlink[\s\S]*?<\/w:hyperlink>/)![0];
      expect(hyperlink).toContain('the spec');
      expect(hyperlink).toContain('<w:rStyle w:val="Hyperlink"/>');

      // Inline order: "See " then the hyperlink then " first"
      const seeIdx = first.indexOf('See ');
      const linkIdx = first.indexOf('<w:hyperlink');
      const firstIdx = first.indexOf(' first');
      expect(seeIdx).toBeGreaterThan(-1);
      expect(linkIdx).toBeGreaterThan(seeIdx);
      expect(firstIdx).toBeGreaterThan(linkIdx);

      // The relationship target part is untouched
      expect(zip.getFileAsString('word/_rels/comments.xml.rels')).toContain(
        'https://example.com/spec'
      );
    } finally {
      doc.dispose();
    }
  });

  it('still exposes hyperlink text through getText() on the parsed comment', async () => {
    const crafted = await buildCraftedDocx();
    const doc = await Document.loadFromBuffer(crafted);
    try {
      const comment = doc.getComment(0)!;
      expect(comment.getText()).toContain('See the spec first');
      expect(comment.getText()).toContain('Second para');
    } finally {
      doc.dispose();
    }
  });
});
