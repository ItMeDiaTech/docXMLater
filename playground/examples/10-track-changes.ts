/**
 * 10 Track Changes: edit an existing document with every change recorded
 * as a Word revision. Open the result in Word to accept or reject edits.
 *
 * Run with: npm run 10-track-changes
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  // 1. Build a starting document.
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

  // 2. Reload it. In real workflows this is a file from someone else.
  const doc = await Document.loadFromBuffer(inputBuffer);

  // 3. Turn on tracked changes. Every modification below becomes a revision.
  doc.enableTrackChanges({ author: 'Reviewer' });

  doc.replaceText('draft margin', 'comfortable margin');

  doc
    .createParagraph('Recommendation: increase the marketing budget for Q3.')
    .setAlignment('justify');

  // 4. Save.
  writeFileSync('10-track-changes.docx', await doc.toBuffer());
  doc.dispose();

  console.log('Wrote 10-track-changes.docx');
  console.log(
    'Open in Word: edits show as struck-through and underlined revisions, attributed to "Reviewer."'
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
