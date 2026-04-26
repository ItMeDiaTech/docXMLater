/**
 * 02 Text: character formatting (bold, italic, color, highlight, sub/superscript).
 *
 * Run with: npm run 02-text
 */

import { Document } from 'docxmlater';
import { writeFileSync } from 'node:fs';

async function main() {
  const doc = Document.create();

  doc.createParagraph('Character Formatting').setStyle('Title');

  const p1 = doc.createParagraph();
  p1.addText('Plain. ');
  p1.addText('Bold. ', { bold: true });
  p1.addText('Italic. ', { italic: true });
  p1.addText('Bold and italic. ', { bold: true, italic: true });
  p1.addText('Underlined. ', { underline: 'single' });
  p1.addText('Strikethrough.', { strike: true });

  const p2 = doc.createParagraph();
  p2.addText('Red. ', { color: 'FF0000' });
  p2.addText('Green. ', { color: '00AA00' });
  p2.addText('Blue. ', { color: '0000FF' });
  p2.addText('Highlighted yellow. ', { highlight: 'yellow' });
  p2.addText('Highlighted cyan.', { highlight: 'cyan' });

  const p3 = doc.createParagraph();
  p3.addText('Chemistry: H');
  p3.addText('2', { subscript: true });
  p3.addText('O. Math: E = mc');
  p3.addText('2', { superscript: true });
  p3.addText('.');

  const p4 = doc.createParagraph();
  p4.addText('Small caps text. ', { smallCaps: true });
  p4.addText('All caps text.', { allCaps: true });

  writeFileSync('02-text.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 02-text.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
