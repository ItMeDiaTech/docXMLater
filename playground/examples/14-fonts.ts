/**
 * 14 Fonts: apply different font families and sizes per run.
 *
 * Run with: npm run 14-fonts
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph('Font Variations').setStyle('Title');

  doc
    .createParagraph()
    .addText('Calibri 11pt: ', { bold: true })
    .addText('The quick brown fox jumps over the lazy dog.', {
      font: 'Calibri',
      size: 11,
    });

  doc
    .createParagraph()
    .addText('Arial 14pt: ', { bold: true })
    .addText('The quick brown fox jumps over the lazy dog.', {
      font: 'Arial',
      size: 14,
    });

  doc
    .createParagraph()
    .addText('Times New Roman 12pt: ', { bold: true })
    .addText('The quick brown fox jumps over the lazy dog.', {
      font: 'Times New Roman',
      size: 12,
    });

  doc.createParagraph().addText('Courier New 10pt: ', { bold: true }).addText('const x = 42;', {
    font: 'Courier New',
    size: 10,
  });

  doc
    .createParagraph()
    .addText('Georgia 13pt: ', { bold: true })
    .addText('The quick brown fox jumps over the lazy dog.', {
      font: 'Georgia',
      size: 13,
    });

  // Document-wide default.
  doc.setDefaultFont('Calibri', 11);

  writeFileSync('14-fonts.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 14-fonts.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
