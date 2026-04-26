/**
 * 09 Bookmarks: mark a section, then jump to it from a hyperlink elsewhere
 * in the document.
 *
 * Run with: npm run 09-bookmarks
 */

import { Document, Hyperlink } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph('Bookmarks').setStyle('Title');

  const target = doc.createBookmark('important_section');

  // Navigation paragraph with a link to the bookmark.
  const nav = doc.createParagraph();
  nav.addText('Click ');
  nav.addHyperlink(
    Hyperlink.createInternal('important_section', 'here', {
      color: '0000FF',
      underline: 'single',
    })
  );
  nav.addText(' to jump to the important section below.');

  // Filler so the jump is visible.
  for (let i = 1; i <= 8; i++) {
    doc.createParagraph(`Filler paragraph ${i}.`);
  }

  // The bookmarked paragraph itself.
  const heading = doc.createParagraph('Important Section').setStyle('Heading1');
  heading.addBookmark(target);

  doc
    .createParagraph(
      'You jumped here by clicking the link above. Bookmarks pair with internal hyperlinks to create in-document navigation.'
    )
    .setAlignment('justify');

  writeFileSync('09-bookmarks.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 09-bookmarks.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
