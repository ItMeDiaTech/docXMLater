import { Document } from './src/core/Document';
import { ZipHandler } from './src/zip/ZipHandler';

async function test() {
  console.log('Loading Original_8.docx...');
  const doc = await Document.load('Original_8.docx');

  // Check field content BEFORE save
  console.log('\n=== Before Save ===');
  checkFieldContent(doc, 'Before');

  // Save document
  console.log('\nSaving to test_regen.docx...');
  await doc.save('test_regen.docx');

  // Check the saved XML
  console.log('\n=== After Save (checking XML) ===');
  const handler = new ZipHandler();
  await handler.load('test_regen.docx');
  const docXml = handler.getFileAsString('word/document.xml') || '';

  const hasBegin = docXml.includes('fldCharType="begin"');
  const hasSeparate = docXml.includes('fldCharType="separate"');
  const hasEnd = docXml.includes('fldCharType="end"');
  const hasInstrText = docXml.includes('instrText');

  console.log('fldChar begin:', hasBegin ? 'YES' : 'NO');
  console.log('fldChar separate:', hasSeparate ? 'YES' : 'NO');
  console.log('fldChar end:', hasEnd ? 'YES' : 'NO');
  console.log('instrText:', hasInstrText ? 'YES' : 'NO');

  // Also check if TOC instruction appears in w:t (bad)
  if (docXml.includes('TOC') && docXml.includes('</w:t>')) {
    // Extract the context around TOC to see if it's in w:t or instrText
    const tocMatch = docXml.match(/<w:[^>]+>[^<]*TOC[^<]*<\/w:[^>]+>/);
    if (tocMatch) {
      console.log('\nTOC context:', tocMatch[0].substring(0, 100));
    }
  }
}

function checkFieldContent(doc: Document, phase: string) {
  const paragraphs = doc.getAllParagraphs();
  let found = 0;

  for (let i = 0; i < paragraphs.length && found < 5; i++) {
    const para = paragraphs[i];
    const runs = para.getRuns();
    for (const run of runs) {
      const content = run.getContent();
      const hasFieldContent = content.some((c: any) =>
        c.type === 'instructionText' || c.type === 'fieldChar'
      );
      if (hasFieldContent) {
        found++;
        console.log(phase + ' - Para ' + i + ': ' + content.map((c: any) => c.type).join(', '));
      }
    }
  }
  console.log('Total field runs found:', found);
}

test().catch(console.error);
