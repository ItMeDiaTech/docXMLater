/**
 * 05 Images: embed an image as a document-level element.
 *
 * For the playground we generate a tiny solid-red 1x1 PNG inline so the
 * example needs no external assets. In a real project you would read the
 * image from disk or fetch it from a URL.
 *
 * Run with: npm run 05-images
 */

import { Document, Image, inchesToEmus } from 'docxmlater';
import { writeFileSync } from 'node:fs';

function createTinyPNG(): Buffer {
  // Smallest valid red 1x1 PNG.
  return Buffer.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
    0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xdd, 0x8d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
    0x44, 0xae, 0x42, 0x60, 0x82,
  ]);
}

async function main() {
  const doc = Document.create();

  doc.createParagraph('Images').setStyle('Title');
  doc.createParagraph(
    'docxmlater accepts PNG, JPEG, GIF, SVG, EMF, and WMF. ' +
      'Images can sit inline with text or float with wrapping.'
  );

  const image = await Image.fromBuffer(createTinyPNG(), {
    width: inchesToEmus(2),
    height: inchesToEmus(2),
    name: 'Sample Image',
    description: 'A 1x1 red pixel scaled to 2 inches square.',
  });
  doc.addImage(image);

  doc.createParagraph('A 1x1 pixel red PNG, scaled to 2 inches square.').setAlignment('center');

  writeFileSync('05-images.docx', await doc.toBuffer());
  doc.dispose();
  console.log('Wrote 05-images.docx');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
