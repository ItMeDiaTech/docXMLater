/**
 * Example: edit an existing document with every change recorded as a Word revision.
 *
 * This is docxmlater's signature feature. Most libraries either cannot edit
 * existing files, or they corrupt the file when revision markup is present.
 *
 * Run with: npm run track-changes
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  // 1. Build the starting document.
  const original = Document.create();
  original.createParagraph().addText('Quarterly Report', {
    bold: true,
    fontSize: 28,
    color: '1F4E79',
  });
  original
    .createParagraph()
    .addText(
      'Revenue grew 12% this quarter. Customer acquisition exceeded our internal forecast by a draft margin.'
    );
  original
    .createParagraph()
    .addText('Operating expenses remained roughly in line with prior quarters.');

  const inputBuffer = await original.toBuffer();
  original.dispose();

  // 2. Reload the document. In real workflows this comes from someone else.
  const doc = await Document.loadFromBuffer(inputBuffer);

  // 3. Turn on tracked changes. Every modification below is now a revision.
  doc.enableTrackChanges({ author: 'Reviewer' });

  // Replace text. Word will show the original struck through and the
  // replacement underlined, both attributed to "Reviewer".
  doc.replaceText('draft margin', 'comfortable margin');

  // Insert a new paragraph as a tracked insertion.
  doc.createParagraph().addText('Recommendation: increase the marketing budget for Q3.', {
    italic: true,
  });

  // 4. Save.
  const outputBuffer = await doc.toBuffer();
  writeFileSync('output-with-revisions.docx', outputBuffer);
  doc.dispose();

  console.log('Wrote output-with-revisions.docx');
  console.log('');
  console.log('In Word you will see:');
  console.log('  - "draft margin" struck through, replaced with "comfortable margin"');
  console.log('  - The new recommendation paragraph underlined and marked as inserted');
  console.log('  - Reviewer attribution on every edit');
  console.log('  - Accept/Reject options under the Review tab');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
