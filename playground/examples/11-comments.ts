/**
 * 11 Comments: attach review comments to paragraphs and resolve some.
 *
 * Run with: npm run 11-comments
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph('Contract Review').setStyle('Title');

  const clause1 = doc.createParagraph('Section 4.2: Payment terms net 30.');
  const clause2 = doc.createParagraph('Section 5.1: Termination requires 90 days written notice.');
  const clause3 = doc.createParagraph('Section 7: Liability cap set at one year of fees paid.');

  // Resolved comment on clause 1.
  const resolved = doc.createComment('Legal', 'Confirmed with finance. Net 30 is acceptable.', 'L');
  clause1.addComment(resolved);
  resolved.resolve();

  // Open comment on clause 2.
  const open1 = doc.createComment(
    'Legal',
    'Should this be 60 days instead? Match precedent from the Globex deal.',
    'L'
  );
  clause2.addComment(open1);

  // Open comment on clause 3.
  const open2 = doc.createComment(
    'Risk',
    'A one-year cap may be too low. Recommend two years.',
    'R'
  );
  clause3.addComment(open2);

  writeFileSync('11-comments.docx', await doc.toBuffer());

  const manager = doc.getCommentManager();
  console.log('Wrote 11-comments.docx');
  console.log(`Open comments:     ${manager.getUnresolvedComments().length}`);
  console.log(`Resolved comments: ${manager.getResolvedComments().length}`);

  doc.dispose();
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
