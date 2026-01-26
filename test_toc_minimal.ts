import { Document } from './src/core/Document';
import { ZipHandler } from './src/zip/ZipHandler';

async function test() {
  console.log('Loading Original_8.docx...');
  const doc = await Document.load('Original_8.docx');
  
  console.log('Saving to test_minimal.docx (no modifications)...');
  await doc.save('test_minimal.docx');
  
  // Extract and check the TOC structure
  const handler = new ZipHandler();
  await handler.load('test_minimal.docx');
  const docXml = handler.getFileAsString('word/document.xml') || '';
  
  const hasBegin = docXml.includes('fldCharType="begin"');
  const hasSeparate = docXml.includes('fldCharType="separate"');
  const hasEnd = docXml.includes('fldCharType="end"');
  const hasInstrText = docXml.includes('instrText');
  const hasTOCInT = docXml.includes('TOC') && docXml.includes('</w:t>');
  
  console.log('\n=== TOC Structure Check ===');
  console.log('fldChar begin:', hasBegin ? 'YES' : 'NO');
  console.log('fldChar separate:', hasSeparate ? 'YES' : 'NO');
  console.log('fldChar end:', hasEnd ? 'YES' : 'NO');
  console.log('instrText:', hasInstrText ? 'YES' : 'NO');
  console.log('TOC in w:t (BAD):', hasTOCInT ? 'YES - BROKEN!' : 'NO - Good');
  
  if (!hasBegin || !hasInstrText) {
    console.log('\nFAILED: TOC structure is broken by docxmlater load/save alone!');
  } else {
    console.log('\nPASSED: TOC structure preserved by docxmlater');
  }
}

test().catch(console.error);
