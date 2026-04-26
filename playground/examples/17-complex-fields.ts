/**
 * 17 Complex Fields: dynamic fields that Word evaluates on open or update.
 *
 * Run with: npm run 17-complex-fields
 */

import { Document, Field } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create({
    properties: {
      title: 'Field Demonstration',
      creator: 'docxmlater',
      subject: 'Dynamic field rendering',
    },
  });

  doc.createParagraph('Dynamic Fields').setStyle('Title');

  // Document properties.
  const props = doc.createParagraph();
  props.addText('Author: ', { bold: true });
  props.addField(Field.createAuthor());

  const titleP = doc.createParagraph();
  titleP.addText('Title: ', { bold: true });
  titleP.addField(Field.createTitle());

  // Date and time.
  const dateP = doc.createParagraph();
  dateP.addText('Date: ', { bold: true });
  dateP.addField(Field.createDate('MMMM d, yyyy'));

  // Page numbering.
  const pageP = doc.createParagraph();
  pageP.addText('This is page ', { bold: true });
  pageP.addField(Field.createPageNumber());
  pageP.addText(' of ');
  pageP.addField(Field.createTotalPages());
  pageP.addText('.');

  // Filename.
  const fileP = doc.createParagraph();
  fileP.addText('Filename: ', { bold: true });
  fileP.addField(Field.createFilename(false));

  doc.createParagraph();
  doc
    .createParagraph(
      'Word evaluates these fields when the document opens. To force a refresh manually, select all and press F9.'
    )
    .setStyle('Normal');

  writeFileSync('17-complex-fields.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 17-complex-fields.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
