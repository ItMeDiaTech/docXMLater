/**
 * 08 Table of Contents: build a document with headings, ready to populate
 * a Word TOC field. Open the result in Word and right-click the TOC, then
 * Update Field, to populate it from the headings below.
 *
 * Run with: npm run 08-table-of-contents
 */

import { Document, TableOfContentsElement } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph('User Guide').setStyle('Title');

  // Insert a TOC at index 1 (after the title). Word renders it as a
  // placeholder until the user updates the field.
  const toc = TableOfContentsElement.createStandard('Table of Contents');
  doc.insertTocAt(1, toc);

  // Headings that Word will collect into the TOC.
  doc.createParagraph('1. Getting Started').setStyle('Heading1');
  doc.createParagraph('Installation').setStyle('Heading2');
  doc.createParagraph('Run npm install docxmlater in your project.');

  doc.createParagraph('First Document').setStyle('Heading2');
  doc.createParagraph('Call Document.create() and then save the result with doc.save().');

  doc.createParagraph('2. Editing Documents').setStyle('Heading1');
  doc.createParagraph('Loading').setStyle('Heading2');
  doc.createParagraph('Document.load(path) reads an existing .docx file.');

  doc.createParagraph('Modifying').setStyle('Heading2');
  doc.createParagraph('Use replaceText, createParagraph, and the element APIs to make changes.');

  doc.createParagraph('3. Advanced Topics').setStyle('Heading1');
  doc.createParagraph(
    'Tracked changes, comments, footnotes, and content controls are all supported.'
  );

  writeFileSync('08-table-of-contents.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 08-table-of-contents.docx');
  console.log('Open in Word, right-click the TOC placeholder, then choose Update Field.');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
