/**
 * Example: attach comments to specific paragraphs and resolve some of them.
 *
 * Comments survive a full save and reload, with their resolution state intact.
 *
 * Run with: npm run comments
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph().addText('Contract Review', {
    bold: true,
    fontSize: 24,
  });

  const clause1 = doc.createParagraph().addText('Section 4.2: Payment terms net 30.');

  const clause2 = doc
    .createParagraph()
    .addText('Section 5.1: Termination requires 90 days written notice.');

  const clause3 = doc
    .createParagraph()
    .addText('Section 7: Liability cap set at one year of fees paid.');

  // Comment on clause 1 - resolved.
  const cleared = doc.createComment('Legal', 'Confirmed with finance. Net 30 is acceptable.', 'L');
  clause1.addComment(cleared);
  cleared.resolve();

  // Comment on clause 2 - still open.
  const open = doc.createComment(
    'Legal',
    'Should this be 60 days instead? Match precedent from the Globex deal.',
    'L'
  );
  clause2.addComment(open);

  // Comment on clause 3 - still open.
  const concern = doc.createComment(
    'Risk',
    'A one-year cap may be too low for the deal size. Recommend two years.',
    'R'
  );
  clause3.addComment(concern);

  const buffer = await doc.toBuffer();
  writeFileSync('contract-review.docx', buffer);

  // Read back the resolution counts before disposing.
  const manager = doc.getCommentManager();
  console.log('Wrote contract-review.docx');
  console.log(`Open comments:     ${manager.getUnresolvedComments().length}`);
  console.log(`Resolved comments: ${manager.getResolvedComments().length}`);

  doc.dispose();
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
