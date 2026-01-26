/**
 * Simulate Template_UI processing steps to identify where corruption occurs
 */

import { Document } from './src/index';

async function analyzeStructure(doc: any, label: string) {
  console.log(`\n--- ${label} ---`);

  const paragraphs = doc.getAllParagraphs();

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

    console.log(`  Paragraph ${i}: ${content.length} items`);

    // Track field structure: begin -> revisions -> separate -> revisions -> end
    let structure = '';
    for (let j = 0; j < content.length; j++) {
      const item = content[j];
      if (!item) continue;
      const typeName = item.constructor.name;

      if (typeName === 'Run') {
        const run = item as any;
        const runContent = run.getContent();
        const fldChar = runContent.find((c: any) => c.type === 'fieldChar');
        if (fldChar) {
          structure += `[${fldChar.fieldCharType}]`;
        } else {
          structure += '[run]';
        }
      } else if (typeName === 'Revision') {
        const revision = item as any;
        structure += `{${revision.getType()}}`;
      } else {
        structure += `<${typeName}>`;
      }
    }
    console.log(`    Structure: ${structure}`);
  }
}

async function testStep(stepName: string, stepFn: () => Promise<void>) {
  console.log(`\n========================================`);
  console.log(`STEP: ${stepName}`);
  console.log(`========================================`);
  await stepFn();
}

async function main() {
  console.log('='.repeat(60));
  console.log('TEMPLATE_UI PROCESSING SIMULATION');
  console.log('='.repeat(60));

  let doc: any;

  // Step 1: Load document with revisions preserved
  await testStep('Load document with revisionHandling: preserve', async () => {
    doc = await Document.load('Original_16.docx', {
      revisionHandling: 'preserve',
      strictParsing: false
    });
    await analyzeStructure(doc, 'After Load');
  });

  // Step 2: Enable track changes (like Template_UI does)
  await testStep('Enable track changes', async () => {
    doc.enableTrackChanges({
      author: 'Test Author',
      trackFormatting: true,
      showInsertionsAndDeletions: true,
    });
    await analyzeStructure(doc, 'After enableTrackChanges');
  });

  // Step 3: Defragment hyperlinks (Template_UI does this)
  await testStep('Defragment hyperlinks', async () => {
    const merged = doc.defragmentHyperlinks({
      resetFormatting: true,
      cleanupRelationships: true,
    });
    console.log(`  Merged ${merged} fragmented hyperlinks`);
    await analyzeStructure(doc, 'After defragmentHyperlinks');
  });

  // Step 4: Standardize hyperlink formatting (Template_UI does this automatically)
  await testStep('Get all paragraphs and iterate (simulating hyperlink processing)', async () => {
    const paragraphs = doc.getAllParagraphs();
    for (const para of paragraphs) {
      const content = para.getContent();
      for (const item of content) {
        // Duck-type check for Hyperlink
        if (item && typeof (item as any).getUrl === 'function') {
          // Just access the hyperlink properties
          const url = (item as any).getUrl();
          const text = (item as any).getText();
        }
        // Check inside revisions
        if (item && typeof (item as any).getContent === 'function') {
          const revContent = (item as any).getContent();
          if (Array.isArray(revContent)) {
            for (const inner of revContent) {
              if (inner && typeof inner.getUrl === 'function') {
                const url = inner.getUrl();
                const text = inner.getText();
              }
            }
          }
        }
      }
    }
    await analyzeStructure(doc, 'After iterating paragraphs');
  });

  // Step 4.5: Simulate creating tracked changes (like URL update with track changes)
  await testStep('Simulate paragraph modifications (add/remove content)', async () => {
    // This simulates what happens when URL updates create tracked changes
    // The issue may be in how paragraph content is modified
    const paragraphs = doc.getAllParagraphs();

    // Access the paragraph content without modifying
    for (const para of paragraphs) {
      const content = para.getContent();
      // Just iterate to ensure content array is accessed
      for (const item of content) {
        if (item && typeof (item as any).getText === 'function') {
          (item as any).getText();
        }
      }
    }
    await analyzeStructure(doc, 'After paragraph content access');
  });

  // Step 4.6: Simulate what applyStyles does - clears formatting
  await testStep('Simulate style application (clear formatting)', async () => {
    const paragraphs = doc.getAllParagraphs();
    for (const para of paragraphs) {
      // Template_UI clears direct formatting when applying styles
      // But only on runs, not on field code runs
      const runs = para.getRuns();
      // Just read runs, don't modify
      for (const run of runs) {
        run.getFormatting();
      }
    }
    await analyzeStructure(doc, 'After style application simulation');
  });

  // Step 5: Save document
  await testStep('Save document', async () => {
    await doc.save('test_simulation_full.docx');
    console.log('  Saved to test_simulation_full.docx');
  });

  // Step 6: Load the saved document and check structure
  await testStep('Load saved document and verify', async () => {
    const doc2 = await Document.load('test_simulation_full.docx', {
      revisionHandling: 'preserve',
      strictParsing: false
    });
    await analyzeStructure(doc2, 'After Load of Saved Document');
    doc2.dispose();
  });

  doc.dispose();

  console.log('\n' + '='.repeat(60));
  console.log('SIMULATION COMPLETE');
  console.log('='.repeat(60));
}

main().catch(console.error);
