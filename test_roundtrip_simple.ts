/**
 * Simple roundtrip test - load and save without any processing
 * This tests if the basic load/save cycle preserves complex field revisions
 */

import { Document } from './src/index';
import * as fs from 'fs';

async function testRoundtrip() {
  console.log('=== Simple Roundtrip Test ===');
  console.log('Loading Original_16.docx...');

  // Load document with revisions preserved
  const doc = await Document.load('Original_16.docx', {
    revisionHandling: 'preserve',
    strictParsing: false
  });

  console.log('Document loaded. Saving immediately as test_roundtrip_output.docx...');

  // Save immediately without any modifications
  await doc.save('test_roundtrip_output.docx');

  console.log('Saved. Now analyzing the output...');

  // Load the output and analyze
  const outputDoc = await Document.load('test_roundtrip_output.docx', {
    revisionHandling: 'preserve',
    strictParsing: false
  });

  const paragraphs = outputDoc.getAllParagraphs();

  // Find paragraphs with HYPERLINK field codes and revisions
  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (!para) continue;

    const content = para.getContent();
    let hasHyperlinkField = false;
    let hasFieldChar = false;

    for (const item of content) {
      if (!item) continue;
      const typeName = item.constructor.name;

      if (typeName === 'Run') {
        const run = item as any;
        const runContent = run.getContent();
        for (const c of runContent) {
          if (c.type === 'fieldChar') hasFieldChar = true;
          if (c.type === 'instructionText' && c.value && c.value.includes('HYPERLINK')) {
            hasHyperlinkField = true;
          }
        }
      } else if (typeName === 'Revision') {
        const revision = item as any;
        const revContent = revision.getContent();
        for (const child of revContent) {
          if (child && child.constructor.name === 'Run') {
            const childRun = child as any;
            const childContent = childRun.getContent();
            for (const c of childContent) {
              if (c.type === 'instructionText' && c.value && c.value.includes('HYPERLINK')) {
                hasHyperlinkField = true;
              }
            }
          }
        }
      }
    }

    if (!hasHyperlinkField || !hasFieldChar) continue;

    console.log(`\n--- Paragraph ${i} (after roundtrip) ---`);
    console.log(`Content items: ${content.length}`);

    for (let j = 0; j < content.length; j++) {
      const item = content[j];
      if (!item) continue;
      const typeName = item.constructor.name;

      if (typeName === 'Run') {
        const run = item as any;
        const runContent = run.getContent();
        const contentTypes = runContent.map((c: any) => {
          if (c.type === 'fieldChar') return `fieldChar(${c.fieldCharType})`;
          if (c.type === 'instructionText') return `instrText`;
          if (c.type === 'text') return `text`;
          return c.type;
        });
        console.log(`  [${j}] Run: [${contentTypes.join(', ')}]`);
      } else if (typeName === 'Revision') {
        const revision = item as any;
        const revType = revision.getType();
        const revId = revision.getId();
        console.log(`  [${j}] Revision: type=${revType}, id=${revId}`);
      } else {
        console.log(`  [${j}] ${typeName}`);
      }
    }
  }

  doc.dispose();
  outputDoc.dispose();

  console.log('\n=== Test Complete ===');
}

testRoundtrip().catch(console.error);
