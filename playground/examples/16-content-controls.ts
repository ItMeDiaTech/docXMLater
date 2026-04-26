/**
 * 16 Content Controls: structured form fields built with Word's Structured
 * Document Tags (SDTs).
 *
 * Run with: npm run 16-content-controls
 */

import { Document, StructuredDocumentTag } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph('New Hire Onboarding Form').setStyle('Title');
  doc.createParagraph(
    'Structured Document Tags (SDTs) are Word’s native form fields. They round-trip cleanly through docxmlater and are accessible to users via the Developer tab in Word.'
  );

  // Plain text input.
  const nameField = StructuredDocumentTag.createPlainText([], false, {
    tag: 'employee_name',
    alias: 'Employee name',
  });
  doc.addStructuredDocumentTag(nameField);

  // Date picker.
  const dateField = StructuredDocumentTag.createDatePicker('M/d/yyyy', [], {
    tag: 'start_date',
    alias: 'Start date',
  });
  doc.addStructuredDocumentTag(dateField);

  // Dropdown list.
  const deptField = StructuredDocumentTag.createDropDownList(
    [
      { displayText: 'Engineering', value: 'eng' },
      { displayText: 'Sales', value: 'sales' },
      { displayText: 'Marketing', value: 'mkt' },
      { displayText: 'Operations', value: 'ops' },
    ],
    [],
    {
      tag: 'department',
      alias: 'Department',
    }
  );
  doc.addStructuredDocumentTag(deptField);

  // Checkbox.
  const policyField = StructuredDocumentTag.createCheckbox(false, [], {
    tag: 'it_policy_accepted',
    alias: 'IT policy accepted',
  });
  doc.addStructuredDocumentTag(policyField);

  writeFileSync('16-content-controls.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 16-content-controls.docx');
  console.log(
    'Open in Word: each form field is a Structured Document Tag. Edit them in place or via the Developer tab.'
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
