import { Document } from './src/core/Document';

async function test() {
  console.log('Loading Original_8.docx...');
  const doc = await Document.load('Original_8.docx');

  // Find TOC paragraphs and check their content
  const paragraphs = doc.getAllParagraphs();

  let foundFieldContent = false;

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const style = para.getStyle() || '';

    const runs = para.getRuns();
    for (let j = 0; j < runs.length; j++) {
      const run = runs[j];
      const content = run.getContent();

      // Check for field content
      const hasFieldContent = content.some((c: any) =>
        c.type === 'instructionText' ||
        c.type === 'fieldChar'
      );

      if (hasFieldContent) {
        foundFieldContent = true;
        console.log('\nPara ' + i + ' (style: ' + (style || 'none') + '), Run ' + j + ':');
        console.log('Content types:', content.map((c: any) => c.type).join(', '));
        for (const c of content) {
          const typedC = c as any;
          if (typedC.type === 'instructionText') {
            console.log('  instructionText:', typedC.value?.substring(0, 50));
          }
          if (typedC.type === 'fieldChar') {
            console.log('  fieldChar:', typedC.fieldCharType);
          }
        }
      }
    }
  }

  if (!foundFieldContent) {
    console.log('\nWARNING: No field content (instructionText/fieldChar) found in any runs!');
    console.log('This means the parser is not preserving field structure.');
  } else {
    console.log('\n=== Field content found - parser is preserving structure ===');
  }
}

test().catch(console.error);
