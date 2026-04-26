/**
 * 03 Lists: bulleted, numbered, and multi-level nested lists.
 *
 * Run with: npm run 03-lists
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph('Lists').setStyle('Title');

  // Bullet list
  doc.createParagraph('Bulleted list:').setStyle('Heading2');
  const bulletId = doc.createBulletList();
  doc.createParagraph('First bullet').setNumbering(bulletId, 0);
  doc.createParagraph('Second bullet').setNumbering(bulletId, 0);
  doc.createParagraph('Third bullet').setNumbering(bulletId, 0);

  // Numbered list
  doc.createParagraph();
  doc.createParagraph('Numbered list:').setStyle('Heading2');
  const numberedId = doc.createNumberedList();
  doc.createParagraph('First step').setNumbering(numberedId, 0);
  doc.createParagraph('Second step').setNumbering(numberedId, 0);
  doc.createParagraph('Third step').setNumbering(numberedId, 0);

  // Multi-level nested list (decimal -> letter -> roman)
  doc.createParagraph();
  doc.createParagraph('Nested list (decimal / letter / roman):').setStyle('Heading2');
  const nestedId = doc.createNumberedList(3, ['decimal', 'lowerLetter', 'lowerRoman']);
  doc.createParagraph('Top-level item one').setNumbering(nestedId, 0);
  doc.createParagraph('Sub-item alpha').setNumbering(nestedId, 1);
  doc.createParagraph('Detail i').setNumbering(nestedId, 2);
  doc.createParagraph('Detail ii').setNumbering(nestedId, 2);
  doc.createParagraph('Sub-item beta').setNumbering(nestedId, 1);
  doc.createParagraph('Top-level item two').setNumbering(nestedId, 0);

  writeFileSync('03-lists.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 03-lists.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
