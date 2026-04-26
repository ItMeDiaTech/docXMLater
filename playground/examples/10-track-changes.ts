/**
 * 10 Track Changes: edit an existing document with every text replacement
 * recorded as a Word revision. Open the result in Word and use the Review
 * tab to accept or reject individual edits.
 *
 * The pattern: enableTrackChanges() then replaceText(). Each replacement
 * becomes a paired <w:del>/<w:ins> revision attributed to the named author.
 *
 * Note: in v11, adding entirely new paragraphs after enableTrackChanges()
 * does not yet emit paragraph-mark insertion markup, which Word can flag as
 * "unreadable content." Stick to text edits within existing paragraphs and
 * the tracked-changes pipeline produces clean, Word-compatible output.
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
    .createParagraph(
      'Operating expenses remained roughly in line with prior quarters.'
    )
    .setAlignment('justify');
  original
    .createParagraph(
      'Recommendation: hold the marketing budget at current levels through Q3.'
    )
    .setAlignment('justify');

  const inputBuffer = await original.toBuffer();
  original.dispose();

  // 2. Reload it. In real workflows this is a file from someone else.
  const doc = await Document.loadFromBuffer(inputBuffer);

  // 3. Turn on tracked changes. Every replaceText below becomes a
  //    paired <w:del>/<w:ins> revision attributed to "Reviewer."
  doc.enableTrackChanges({ author: 'Reviewer' });

  doc.replaceText('draft margin', 'comfortable margin');
  doc.replaceText(
    'hold the marketing budget at current levels',
    'increase the marketing budget by 15%'
  );

  // 4. Save.
  writeFileSync('10-track-changes.docx', await doc.toBuffer());
  doc.dispose();

  console.log('Wrote 10-track-changes.docx');
  console.log('');
  console.log('Open in Word and you will see:');
  console.log(
    '  - "draft margin" struck through, "comfortable margin" underlined'
  );
  console.log(
    '  - "hold...current levels" struck through, "increase...by 15%" underlined'
  );
  console.log('  - Both edits attributed to "Reviewer" in the Review tab');
  console.log(
    '  - Accept All / Reject All work cleanly with no recovery prompts'
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
