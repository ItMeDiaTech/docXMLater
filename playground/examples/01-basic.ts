/**
 * 01 Basic: the smallest possible docxmlater program.
 *
 * Run with: npm run 01-basic
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();
  doc.createParagraph('Hello, World!').setStyle('Title');
  doc.createParagraph('This is a minimal document. It has a title and one body paragraph.');

  writeFileSync('01-basic.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 01-basic.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
