/**
 * normalizeSpacing({ removeDuplicateEmptyParagraphs }) must not delete
 * text-empty paragraphs that still carry content.
 *
 * Emptiness was judged solely by getText().trim() === '', but a paragraph with
 * no visible text can still hold a bookmark anchor (REF/hyperlink target), an
 * inline image (w:drawing), a section break (sectPr), a range marker, or a
 * comment anchor. Splicing such a paragraph out silently destroys that content.
 * The fix mirrors TableCell.isParaBlank: only paragraphs that pass the richer
 * blankness check are eligible for removal as duplicate empties.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Bookmark } from '../../src/elements/Bookmark';
import { Image } from '../../src/elements/Image';
import { ImageRun } from '../../src/elements/ImageRun';

function countImageParagraphs(doc: Document): number {
  let count = 0;
  for (const para of doc.getParagraphs()) {
    if (para.getContent().some((item) => item instanceof ImageRun)) count++;
  }
  return count;
}

const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

describe('normalizeSpacing preserves content-bearing empty paragraphs', () => {
  it('keeps an empty paragraph carrying a bookmark even after another empty paragraph', async () => {
    const doc = Document.create();

    doc.addParagraph(new Paragraph().addText('Heading'));
    // First empty paragraph (truly blank).
    doc.addParagraph(new Paragraph());
    // Second empty paragraph carries a bookmark — must NOT be removed.
    const anchorPara = new Paragraph();
    anchorPara.addBookmark(new Bookmark({ name: 'TargetAnchor' }));
    doc.addParagraph(anchorPara);

    expect(doc.getBookmarks().length).toBe(1);

    doc.normalizeSpacing({ removeDuplicateEmptyParagraphs: true });

    // The bookmark anchor survives — otherwise REF fields / internal hyperlinks
    // targeting it would break.
    const bookmarks = doc.getBookmarks();
    expect(bookmarks.length).toBe(1);
    expect(bookmarks[0]!.bookmark.getName()).toBe('TargetAnchor');

    doc.dispose();
  });

  it('keeps an image-only paragraph that follows an empty paragraph', async () => {
    const doc = Document.create();

    doc.addParagraph(new Paragraph().addText('Heading'));
    doc.addParagraph(new Paragraph());
    // addImage appends an image-only paragraph: getText() is '', but it carries
    // a w:drawing — it must not be classified as a removable duplicate empty.
    const image = await Image.fromBuffer(Buffer.from(PNG_BASE64, 'base64'), {
      width: 914400,
      height: 914400,
    });
    doc.addImage(image);

    expect(countImageParagraphs(doc)).toBe(1);

    doc.normalizeSpacing({ removeDuplicateEmptyParagraphs: true });

    // The image-bearing paragraph must still be present in the body.
    expect(countImageParagraphs(doc)).toBe(1);

    doc.dispose();
  });

  it('still removes genuinely blank duplicate paragraphs', async () => {
    const doc = Document.create();

    doc.addParagraph(new Paragraph().addText('Heading'));
    doc.addParagraph(new Paragraph());
    doc.addParagraph(new Paragraph());
    doc.addParagraph(new Paragraph().addText('Body'));

    const result = doc.normalizeSpacing({ removeDuplicateEmptyParagraphs: true });
    expect(result.removed).toBe(1);

    doc.dispose();
  });
});
