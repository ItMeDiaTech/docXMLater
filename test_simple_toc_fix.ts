import { Document } from './src/core/Document';
import { ZipHandler } from './src/zip/ZipHandler';

async function test() {
  console.log('Loading Original_8.docx...');
  const doc = await Document.load('Original_8.docx');

  // Enable auto-populate TOCs (this is what Template_UI does)
  console.log('Enabling setAutoPopulateTOCs(true)...');
  doc.setAutoPopulateTOCs(true);

  // Save document
  console.log('Saving to test_simple_toc_fix.docx...');
  await doc.save('test_simple_toc_fix.docx');

  // Check the saved XML
  console.log('\n=== Checking saved XML ===');
  const handler = new ZipHandler();
  await handler.load('test_simple_toc_fix.docx');
  const docXml = handler.getFileAsString('word/document.xml') || '';

  const hasBegin = docXml.includes('fldCharType="begin"');
  const hasSeparate = docXml.includes('fldCharType="separate"');
  const hasEnd = docXml.includes('fldCharType="end"');
  const hasInstrText = docXml.includes('instrText');

  console.log('fldChar begin:', hasBegin ? 'YES' : 'NO');
  console.log('fldChar separate:', hasSeparate ? 'YES' : 'NO');
  console.log('fldChar end:', hasEnd ? 'YES' : 'NO');
  console.log('instrText:', hasInstrText ? 'YES' : 'NO');

  // Check if TOC instruction is in instrText (good) or w:t (bad)
  const tocInInstrText = docXml.match(/<w:instrText[^>]*>[^<]*TOC[^<]*<\/w:instrText>/);
  const tocInT = docXml.match(/<w:t[^>]*>[^<]*TOC[^<]*\\[^<]*<\/w:t>/);

  console.log('\nTOC instruction in instrText (GOOD):', tocInInstrText ? 'YES' : 'NO');
  console.log('TOC instruction in w:t (BAD):', tocInT ? 'YES' : 'NO');

  if (hasBegin && hasSeparate && hasEnd && hasInstrText && tocInInstrText && !tocInT) {
    console.log('\n=== SUCCESS: TOC field structure is correct! ===');
  } else {
    console.log('\n=== FAILED: TOC field structure is broken ===');
  }
}

test().catch(console.error);
