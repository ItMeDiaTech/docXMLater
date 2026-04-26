/**
 * 06 Headers and Footers: title in the header, page numbering in the footer.
 *
 * Run with: npm run 06-headers-footers
 */

import { Document, Header, Footer, Field } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  // Header: title aligned right.
  const header = Header.createDefault();
  header.createParagraph().setAlignment('right').addText('Quarterly Report', {
    bold: true,
  });
  doc.setHeader(header);

  // Footer: "Page X of Y" centered.
  const footer = Footer.createDefault();
  const f = footer.createParagraph().setAlignment('center');
  f.addText('Page ');
  f.addField(Field.createPageNumber());
  f.addText(' of ');
  f.addField(Field.createTotalPages());
  doc.setFooter(footer);

  // Body content - several paragraphs to make pagination meaningful.
  doc.createParagraph('Quarterly Report').setStyle('Title');
  for (let i = 1; i <= 4; i++) {
    doc.createParagraph(`Section ${i}`).setStyle('Heading1');
    doc
      .createParagraph('Lorem ipsum dolor sit amet, consectetur adipiscing elit. '.repeat(8))
      .setAlignment('justify');
  }

  writeFileSync('06-headers-footers.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 06-headers-footers.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
