import { Document } from './src/core/Document';
import { ZipHandler } from './src/zip/ZipHandler';

async function test() {
  console.log('Loading Original_8.docx...');
  const doc = await Document.load('Original_8.docx');

  // Check initial state
  console.log('\n=== Initial State ===');
  checkFieldContent(doc);

  // Simulate Template_UI processing steps one by one

  // Step 1: Enable track changes (like Template_UI does)
  console.log('\n=== After enableTrackChanges ===');
  doc.enableTrackChanges({ author: 'Test' });
  checkFieldContent(doc);

  // Step 2: Iterate through paragraphs and runs (like removeExtraWhitespace does)
  console.log('\n=== After iterating paragraphs/runs ===');
  const paragraphs = doc.getAllParagraphs();
  for (const para of paragraphs) {
    const runs = para.getRuns();
    for (const run of runs) {
      // Just access the run text (don't modify)
      const text = run.getText();
    }
  }
  checkFieldContent(doc);

  // Step 3: Call setFont on a regular paragraph (not TOC)
  console.log('\n=== After setFont on non-TOC paragraph ===');
  const firstNonTocPara = paragraphs.find(p => !p.getStyle()?.startsWith('TOC'));
  if (firstNonTocPara) {
    const runs = firstNonTocPara.getRuns();
    if (runs.length > 0) {
      runs[0].setFont('Verdana');
    }
  }
  checkFieldContent(doc);

  // Step 4: Call setAutoPopulateTOCs (like Template_UI does before save)
  console.log('\n=== After setAutoPopulateTOCs(true) ===');
  doc.setAutoPopulateTOCs(true);
  checkFieldContent(doc);

  // Step 5: Save
  console.log('\n=== Saving ===');
  await doc.save('test_template_sim.docx');

  // Check saved XML
  console.log('\n=== Checking saved XML ===');
  const handler = new ZipHandler();
  await handler.load('test_template_sim.docx');
  const docXml = handler.getFileAsString('word/document.xml') || '';

  const hasBegin = docXml.includes('fldCharType="begin"');
  const hasInstrText = docXml.includes('instrText');
  const hasTOCInT = docXml.match(/<w:t[^>]*>[^<]*TOC[^<]*\\[^<]*<\/w:t>/);

  console.log('fldChar begin:', hasBegin ? 'YES' : 'NO');
  console.log('instrText:', hasInstrText ? 'YES' : 'NO');
  console.log('TOC in w:t (BAD):', hasTOCInT ? 'YES - BROKEN!' : 'NO - Good');
}

function checkFieldContent(doc: Document) {
  const paragraphs = doc.getAllParagraphs();
  let found = 0;

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const runs = para.getRuns();
    for (const run of runs) {
      const content = run.getContent();
      const hasFieldContent = content.some((c: any) =>
        c.type === 'instructionText' || c.type === 'fieldChar'
      );
      if (hasFieldContent) {
        found++;
      }
    }
  }
  console.log('Field runs found:', found);
}

test().catch(console.error);
