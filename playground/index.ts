/**
 * docxmlater Playground - main entry point
 *
 * This is the default example. It demonstrates docxmlater's signature feature:
 * editing a document while tracking every change as a Word revision.
 *
 * Steps:
 *   1. Generate a starting document.
 *   2. Reload it to simulate "incoming" content from someone else.
 *   3. Apply tracked-change edits as the "Reviewer" author.
 *   4. Save the result. Open it in Word and you'll see real revision marks.
 *
 * Try editing this file, then run `npm start` in the terminal panel.
 *
 * Twenty topical examples are also available - see README.md for the full
 * list, or run any of `npm run 01-basic` through `npm run 20-compatibility-mode`.
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  // 1. Build a starting document and save it as our "input."
  const original = Document.create();
  original.createParagraph().addText('Quarterly Report', {
    bold: true,
    size: 28,
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
  writeFileSync('input.docx', inputBuffer);
  console.log('Wrote input.docx');

  // 2. Reload the document. In real workflows this would be a file from
  //    somewhere else (email, shared drive, a database blob).
  const doc = await Document.loadFromBuffer(inputBuffer);

  // 3. Enable tracked changes and apply edits as the "Reviewer."
  doc.enableTrackChanges({ author: 'Reviewer' });

  // Replace a phrase. The replacement is recorded as a tracked change.
  doc.replaceText('draft margin', 'comfortable margin');

  // Insert a new paragraph - also recorded as a tracked insertion.
  doc.createParagraph().addText('Recommendation: increase the marketing budget for Q3.', {
    italic: true,
  });

  // 4. Save the result.
  const outputBuffer = await doc.toBuffer();
  writeFileSync('output.docx', outputBuffer);
  doc.dispose();

  console.log('Wrote output.docx');
  console.log('');
  console.log('Open output.docx in Word and you will see:');
  console.log('  - Strikethrough on "draft margin"');
  console.log('  - Underlined insertion of "comfortable margin"');
  console.log('  - The new recommendation paragraph as an inserted change');
  console.log('  - All edits attributed to "Reviewer"');
  console.log('');
  console.log('In StackBlitz, right-click output.docx in the Files panel and choose Download.');
}

main().catch((err) => {
  console.error('Playground error:', err);
  process.exit(1);
});
