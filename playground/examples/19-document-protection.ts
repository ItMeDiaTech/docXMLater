/**
 * 19 Document Protection: restrict edits to tracked changes only.
 *
 * Run with: npm run 19-document-protection
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph('Protected Contract Draft').setStyle('Title');

  doc.createParagraph(
    'This document is locked in tracked-changes mode. Any modifications a reviewer makes in Word will be recorded as revisions and can be accepted or rejected before publication.'
  );

  doc.createParagraph('Section 1: Term').setStyle('Heading1');
  doc.createParagraph(
    'This Agreement begins on the Effective Date and continues for an initial term of three (3) years.'
  );

  doc.createParagraph('Section 2: Payment').setStyle('Heading1');
  doc.createParagraph('All invoices are payable net thirty (30) days from receipt.');

  // Lock the document into tracked-changes-only editing.
  // (Password protection is intentionally omitted; the playground runs in an
  // ESM context where v11's password hashing path is unavailable.)
  doc.protectDocument({
    edit: 'trackedChanges',
  });

  writeFileSync('19-document-protection.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 19-document-protection.docx');
  console.log(
    'In Word: editing is restricted to tracked changes only. Use the Review tab to stop protection.'
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
