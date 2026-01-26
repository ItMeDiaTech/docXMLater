/**
 * Diagnostic script to analyze paragraph content structure for complex field revisions
 * This helps identify where the bug is in the parsing/regeneration pipeline
 */

import { Document } from './src/index';

async function analyzeDocument(filePath: string, label: string) {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`Analyzing: ${label}`);
  console.log(`File: ${filePath}`);
  console.log('='.repeat(60));

  const doc = await Document.load(filePath, {
    revisionHandling: 'preserve',
    strictParsing: false
  });

  const paragraphs = doc.getAllParagraphs();

  // Find paragraphs containing complex fields or revisions
  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (!para) continue;

    const content = para.getContent();

    // Check if this paragraph has field characters or revisions
    let hasFieldChar = false;
    let hasRevision = false;
    let hasHyperlinkField = false;

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
        hasRevision = true;
        // Check inside revision for field content
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

    // Only show paragraphs with HYPERLINK fields and revisions
    if (!hasHyperlinkField) continue;

    const text = para.getText() || '';
    console.log(`\n--- Paragraph ${i} ---`);
    console.log(`Text: "${text.substring(0, 100)}${text.length > 100 ? '...' : ''}"`);
    console.log(`Has field chars: ${hasFieldChar}, Has revisions: ${hasRevision}, Has HYPERLINK: ${hasHyperlinkField}`);
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
          if (c.type === 'instructionText') return `instrText(${c.isDeleted ? 'DEL' : ''}): "${(c.value || '').substring(0, 40)}..."`;
          if (c.type === 'text') return `text(${c.isDeleted ? 'DEL' : ''}): "${(c.value || '').substring(0, 40)}..."`;
          return c.type;
        });
        console.log(`  [${j}] Run: [${contentTypes.join(', ')}]`);
      } else if (typeName === 'Revision') {
        const revision = item as any;
        const revType = revision.getType();
        const revId = revision.getId();
        const revAuthor = revision.getAuthor() || '';
        const revContent = revision.getContent();
        console.log(`  [${j}] Revision: type=${revType}, id=${revId}, author="${revAuthor.substring(0, 20)}"`);

        // Show content inside revision
        for (let k = 0; k < revContent.length; k++) {
          const child = revContent[k];
          if (!child) continue;
          const childType = child.constructor.name;
          if (childType === 'Run') {
            const childRun = child as any;
            const childContent = childRun.getContent();
            const childTypes = childContent.map((c: any) => {
              if (c.type === 'fieldChar') return `fieldChar(${c.fieldCharType})`;
              if (c.type === 'instructionText') return `instrText: "${(c.value || '').substring(0, 40)}..."`;
              if (c.type === 'text') return `text: "${(c.value || '').substring(0, 40)}..."`;
              return c.type;
            });
            console.log(`       [${k}] Run: [${childTypes.join(', ')}]`);
          } else {
            console.log(`       [${k}] ${childType}`);
          }
        }
      } else if (typeName === 'Hyperlink') {
        const hyperlink = item as any;
        const url = hyperlink.getUrl() || hyperlink.getAnchor() || '';
        const hlText = hyperlink.getText() || '';
        console.log(`  [${j}] Hyperlink: url="${url.substring(0, 50)}...", text="${hlText.substring(0, 40)}..."`);
      } else if (typeName === 'ComplexField') {
        const field = item as any;
        const instr = field.getInstruction() || '';
        console.log(`  [${j}] ComplexField: instruction="${instr.substring(0, 50)}..."`);

        // Check for result revisions
        if (field.hasResultRevisions && field.hasResultRevisions()) {
          const resultRevisions = field.getResultRevisions();
          console.log(`       Result revisions: ${resultRevisions.length}`);
          for (const rev of resultRevisions) {
            console.log(`         - type=${rev.getType()}, id=${rev.getId()}`);
          }
        }
      } else {
        console.log(`  [${j}] ${typeName}`);
      }
    }
  }

  doc.dispose();
}

async function main() {
  try {
    // Analyze original document
    await analyzeDocument('Original_16.docx', 'ORIGINAL');

    // Analyze processed document
    await analyzeDocument('Processed_16.docx', 'PROCESSED');

    console.log('\n' + '='.repeat(60));
    console.log('Analysis complete');
    console.log('='.repeat(60));
  } catch (error) {
    console.error('Error:', error);
  }
}

main();
