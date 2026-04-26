/**
 * docxmlater Playground - main entry point
 *
 * The default example demonstrates the library's signature workflow:
 * load an existing .docx, modify it, and save it back without breaking
 * anything. The output document opens cleanly in Word with no
 * "unreadable content" warnings.
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
  original.createParagraph('Quarterly Report').setStyle('Title');
  original
    .createParagraph(
      'Revenue grew 12% this quarter. Customer acquisition exceeded our internal forecast by a draft margin.'
    )
    .setAlignment('justify');
  original
    .createParagraph('Operating expenses remained roughly in line with prior quarters.')
    .setAlignment('justify');

  const inputBuffer = await original.toBuffer();
  original.dispose();
  writeFileSync('input.docx', inputBuffer);
  console.log('Wrote input.docx');

  // 2. Reload the document. In real workflows this is a file from
  //    somewhere else - email, a shared drive, a database blob.
  const doc = await Document.loadFromBuffer(inputBuffer);

  // 3. Modify it. The original styling, page setup, and metadata are
  //    preserved through the round-trip.
  doc.replaceText('draft margin', 'comfortable margin');
  doc.findAndHighlight('comfortable margin', 'yellow');

  doc.createParagraph('Recommendation').setStyle('Heading2');
  doc
    .createParagraph(
      'Increase the marketing budget for Q3 by 15% to capitalize on the favorable margin.'
    )
    .setAlignment('justify');

  // 4. Save. The output is a clean, schema-valid .docx that opens in Word
  //    without any recovery prompts.
  writeFileSync('output.docx', await doc.toBuffer());
  doc.dispose();

  console.log('Wrote output.docx');
  console.log('');
  console.log('Open output.docx in Word and you will see:');
  console.log('  - "draft margin" replaced with a yellow-highlighted "comfortable margin"');
  console.log('  - A new "Recommendation" heading and follow-up paragraph');
  console.log('  - All original styling preserved through the round-trip');
  console.log('');
  console.log('For a tracked-changes demo, run: npm run 10-track-changes');
  console.log('In StackBlitz, right-click output.docx in the Files panel and choose Download.');
}

main().catch((err) => {
  console.error('Playground error:', err);
  process.exit(1);
});
