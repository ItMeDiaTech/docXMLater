/**
 * Simple load and save example using docxmlater
 * Loads test-document-output.docx and saves it as output.docx
 */

import { Document } from './src/index';

async function main() {
  try {
    console.log('Loading test-document-output.docx...');
    const doc = await Document.load('test-document-output.docx');

    console.log('Saving to output.docx...');
    await doc.save('output.docx');

    console.log('✓ Successfully loaded and saved document');
  } catch (error) {
    console.error('Error:', error);
    process.exit(1);
  }
}

// Run if executed directly
if (require.main === module) {
  main();
}

export { main };
