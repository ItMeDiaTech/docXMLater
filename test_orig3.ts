import { Document } from './src';
import * as fs from 'fs';

async function test() {
  console.log('Testing Corruption_Original_3.docx with docxmlater directly...\n');

  const JSZip = require('jszip');

  async function countBookmarks(filePath: string) {
    const buffer = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(buffer);
    const docXml = await zip.file('word/document.xml').async('string');
    const startCount = (docXml.match(/<w:bookmarkStart/g) || []).length;
    const endCount = (docXml.match(/<w:bookmarkEnd/g) || []).length;
    return { startCount, endCount };
  }

  const original = await countBookmarks('Corruption_Original_3.docx');
  console.log('Original: ' + original.startCount + ' starts, ' + original.endCount + ' ends');

  // Load and save with docxmlater
  const doc = await Document.load('Corruption_Original_3.docx');
  await doc.save('Corruption_Test_3.docx');
  doc.dispose();

  const processed = await countBookmarks('Corruption_Test_3.docx');
  console.log('After docxmlater: ' + processed.startCount + ' starts, ' + processed.endCount + ' ends');

  if (processed.startCount === original.startCount && processed.endCount === original.endCount) {
    console.log('\nSUCCESS: docxmlater preserves all bookmarks');
  } else {
    console.log('\nFAILURE: docxmlater loses bookmarks');
    console.log('  Missing starts: ' + (original.startCount - processed.startCount));
    console.log('  Missing ends: ' + (original.endCount - processed.endCount));
  }

  fs.unlinkSync('Corruption_Test_3.docx');
}

test().catch(console.error);
