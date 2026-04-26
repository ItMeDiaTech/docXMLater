/**
 * 04 Styles: built-in styles plus a custom style applied to a paragraph.
 *
 * Run with: npm run 04-styles
 */

import { Document, Style } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  // Built-in styles.
  doc.createParagraph('Built-in styles').setStyle('Title');
  doc.createParagraph('A subtitle in the Subtitle style').setStyle('Subtitle');
  doc.createParagraph('Heading 1').setStyle('Heading1');
  doc.createParagraph('Heading 2').setStyle('Heading2');
  doc.createParagraph('Heading 3').setStyle('Heading3');
  doc
    .createParagraph(
      'Body text in the Normal style. Built-in styles include Title, Subtitle, ' +
        'Heading1 through Heading9, Normal, and ListParagraph.'
    )
    .setStyle('Normal');

  // Custom style.
  doc.createParagraph();
  doc.createParagraph('Custom style').setStyle('Heading1');

  const callout = new Style({
    styleId: 'Callout',
    name: 'Callout',
    type: 'paragraph',
    customStyle: true,
    runFormatting: {
      bold: true,
      size: 14,
      color: 'C00000',
    },
    paragraphFormatting: {
      alignment: 'center',
      spacing: { before: 240, after: 240 },
    },
  });
  doc.getStylesManager().addStyle(callout);

  doc
    .createParagraph('This paragraph uses the custom Callout style defined above.')
    .setStyle('Callout');

  writeFileSync('04-styles.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 04-styles.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
