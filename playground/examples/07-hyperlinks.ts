/**
 * 07 Hyperlinks: external web link, email link, and a custom-styled link.
 *
 * Run with: npm run 07-hyperlinks
 */

import { Document, Hyperlink } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph('Hyperlinks').setStyle('Title');

  // External web link.
  const p1 = doc.createParagraph();
  p1.addText('Project repository: ');
  p1.addHyperlink(
    Hyperlink.createWebLink('https://github.com/ItMeDiaTech/docXMLater', 'docxmlater on GitHub')
  );

  // Email link.
  const p2 = doc.createParagraph();
  p2.addText('Send feedback to: ');
  p2.addHyperlink(Hyperlink.createEmail('issues@example.com'));

  // Custom-styled link.
  const p3 = doc.createParagraph();
  p3.addText('Styled link: ');
  p3.addHyperlink(
    Hyperlink.createWebLink('https://example.com', 'big bold red link', {
      bold: true,
      color: 'C00000',
      size: 14,
    })
  );

  // Multiple links inline in one sentence.
  const p4 = doc.createParagraph();
  p4.setAlignment('justify');
  p4.addText('See ');
  p4.addHyperlink(Hyperlink.createWebLink('https://github.com', 'GitHub'));
  p4.addText(' for code, ');
  p4.addHyperlink(Hyperlink.createWebLink('https://stackoverflow.com', 'Stack Overflow'));
  p4.addText(' for Q&A, and ');
  p4.addHyperlink(Hyperlink.createWebLink('https://www.typescriptlang.org', 'TypeScript'));
  p4.addText(' for the language documentation.');

  writeFileSync('07-hyperlinks.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 07-hyperlinks.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
