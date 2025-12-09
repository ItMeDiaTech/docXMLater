/**
 * Debug script to check why images aren't being processed
 */
import { Document } from './src/core/Document';
import { ImageRun } from './src/elements/ImageRun';
import { Revision } from './src/elements/Revision';

async function debugImages() {
  console.log('=== Loading Example1_Original.docx ===\n');

  const doc = await Document.load('./Example1_Original.docx');

  const paragraphs = doc.getAllParagraphs();
  console.log(`Total paragraphs found: ${paragraphs.length}\n`);

  let imageRunCount = 0;
  let largeImageCount = 0;
  const minEmus = 96 * 9525; // 1 inch at 96 DPI

  console.log('=== Scanning paragraphs for images ===\n');

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const content = para.getContent();

    for (const item of content) {
      // Check the type of each content item
      const typeName = item.constructor.name;

      if (item instanceof ImageRun) {
        imageRunCount++;
        const image = item.getImageElement();
        const width = image.getWidth();
        const height = image.getHeight();
        const isLarge = width >= minEmus || height >= minEmus;

        if (isLarge) largeImageCount++;

        console.log(`Paragraph ${i}: Found ImageRun`);
        console.log(`  - Type: ${typeName}`);
        console.log(`  - Size: ${(width / 914400).toFixed(2)}" x ${(height / 914400).toFixed(2)}"`);
        console.log(`  - Is Large (>=1"): ${isLarge}`);
        console.log('');
      } else if (item instanceof Revision) {
        // Check revision content
        const revContent = item.getContent();
        for (const revItem of revContent) {
          if (revItem instanceof ImageRun) {
            imageRunCount++;
            const image = revItem.getImageElement();
            const width = image.getWidth();
            const height = image.getHeight();
            const isLarge = width >= minEmus || height >= minEmus;

            if (isLarge) largeImageCount++;

            console.log(`Paragraph ${i}: Found ImageRun inside Revision`);
            console.log(`  - Revision Type: ${item.getType()}`);
            console.log(`  - Size: ${(width / 914400).toFixed(2)}" x ${(height / 914400).toFixed(2)}"`);
            console.log(`  - Is Large (>=1"): ${isLarge}`);
            console.log('');
          }
        }
      } else {
        // Log other types to see what's there
        if (typeName !== 'Run') {
          // console.log(`Paragraph ${i}: Found ${typeName}`);
        }
      }
    }
  }

  console.log('=== Summary ===');
  console.log(`Total ImageRun found: ${imageRunCount}`);
  console.log(`Large images (>=1"): ${largeImageCount}`);

  console.log('\n=== Calling borderAndCenterLargeImages() ===');
  const result = doc.borderAndCenterLargeImages();
  console.log(`Images bordered and centered: ${result}`);

  // Save to check the output
  await doc.save('./Example1_Debug_Output.docx');
  console.log('\nSaved to Example1_Debug_Output.docx');
}

debugImages().catch(console.error);
