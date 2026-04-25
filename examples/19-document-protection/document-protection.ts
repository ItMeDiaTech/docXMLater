/**
 * Example: Document protection.
 *
 * Demonstrates the four enforcement modes Word supports:
 *   - readOnly      — view only; no edits allowed
 *   - comments      — only comments may be added
 *   - trackedChanges — edits are allowed but force tracked changes
 *   - forms         — only form-field controls may be filled in
 *
 * Optional password protection generates a salted PBKDF2 hash that
 * Word will require to unlock the protection state.
 *
 * Run: `npx ts-node examples/19-document-protection/document-protection.ts`
 */

import { Document } from '../../src';

async function main() {
  // 1. Read-only document with no password
  const readOnly = Document.create();
  readOnly.createParagraph('Read-only report').setStyle('Heading1');
  readOnly.createParagraph('This document cannot be edited in Word.');
  readOnly.protectDocument({ edit: 'readOnly', enforcement: true });
  await readOnly.save('examples/19-document-protection/readonly.docx');
  readOnly.dispose();

  // 2. Tracked-changes-required document (edits allowed, all auto-revisioned)
  const trackedOnly = Document.create();
  trackedOnly.createParagraph('Collaborative draft').setStyle('Heading1');
  trackedOnly.createParagraph('Any edit will be recorded as a tracked change.');
  trackedOnly.protectDocument({ edit: 'trackedChanges', enforcement: true });
  await trackedOnly.save('examples/19-document-protection/tracked-only.docx');
  trackedOnly.dispose();

  // 3. Comments-only with password
  const commentsOnly = Document.create();
  commentsOnly.createParagraph('Comments-only review').setStyle('Heading1');
  commentsOnly.createParagraph('Reviewers can add comments but cannot edit text.');
  commentsOnly.protectDocument({
    edit: 'comments',
    enforcement: true,
    password: 'reviewer-secret',
    cryptSpinCount: 100_000,
  });
  await commentsOnly.save('examples/19-document-protection/comments-only.docx');
  commentsOnly.dispose();

  // 4. Forms protection (only form fields fillable)
  const forms = Document.create();
  forms.createParagraph('Application form').setStyle('Heading1');
  forms.createParagraph('Fill out the fields below.');
  forms.protectDocument({ edit: 'forms', enforcement: true });
  await forms.save('examples/19-document-protection/forms.docx');
  forms.dispose();

  // eslint-disable-next-line no-console
  console.log('Wrote 4 protected documents under examples/19-document-protection/');
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error(err);
  process.exit(1);
});
