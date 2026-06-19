/**
 * replaceText() is a literal-string API (the find side is regex-escaped), so the
 * replacement must be inserted verbatim. Previously the replacement was passed
 * straight to String.prototype.replace, where $&, $', $`, and $$ are
 * substitution tokens — so a caller-supplied replacement containing those
 * sequences produced silently wrong document text. Using a replacer function
 * inserts the string literally.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';

describe('replaceText inserts $-tokens literally', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('does not expand $& / $‘ substitution tokens in the replacement', () => {
    doc = Document.create();
    const para = new Paragraph();
    para.addText('Total: X end');
    doc.addParagraph(para);

    doc.replaceText('X', "$&$'");

    expect(para.getText()).toBe("Total: $&$' end");
  });

  it('does not halve $$ in the replacement', () => {
    doc = Document.create();
    const para = new Paragraph();
    para.addText('Price: X');
    doc.addParagraph(para);

    doc.replaceText('X', '$$100');

    expect(para.getText()).toBe('Price: $$100');
  });

  it('does not expand $-tokens in the whole-word branch', () => {
    doc = Document.create();
    const para = new Paragraph();
    para.addText('cost X total');
    doc.addParagraph(para);

    doc.replaceText('X', '$$50', { wholeWord: true });

    expect(para.getText()).toBe('cost $$50 total');
  });
});
