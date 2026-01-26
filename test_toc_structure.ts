import { Document } from './src/core/Document';

async function test() {
  console.log('Loading Original_8.docx...');
  const doc = await Document.load('Original_8.docx');

  console.log('\n=== Checking TOC Field Structure After Load ===');
  const paragraphs = doc.getAllParagraphs();
  
  let tocFieldFound = false;
  let hasInstrText = false;
  let hasFieldChar = false;
  
  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i]!;
    const runs = para.getRuns();
    
    for (const run of runs) {
      const content = run.getContent();
      for (const c of content) {
        if ((c as any).type === 'instructionText') {
          const val = (c as any).value || '';
          if (val.includes('TOC')) {
            console.log(`Found instructionText with TOC at paragraph ${i}: "${val.substring(0, 50)}..."`);
            tocFieldFound = true;
            hasInstrText = true;
          }
        }
        if ((c as any).type === 'fieldChar') {
          const charType = (c as any).fieldCharType;
          console.log(`Found fieldChar (${charType}) at paragraph ${i}`);
          hasFieldChar = true;
        }
      }
    }
  }
  
  console.log('\n=== Summary ===');
  console.log('TOC field found:', tocFieldFound ? 'YES' : 'NO');
  console.log('Has instructionText:', hasInstrText ? 'YES' : 'NO');
  console.log('Has fieldChar:', hasFieldChar ? 'YES' : 'NO');
  
  if (tocFieldFound && hasInstrText && hasFieldChar) {
    console.log('\nSUCCESS: TOC field structure is preserved after parsing!');
  } else {
    console.log('\nFAILED: TOC field structure was not preserved!');
  }
}

test().catch(console.error);
