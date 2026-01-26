/**
 * Full Template_UI processing simulation
 * Closely matches the actual processing flow to identify where corruption occurs
 */

import { Document } from './src/index';
import * as fs from 'fs';
import JSZip from 'jszip';

async function analyzeXml(filePath: string, label: string) {
  console.log(`\n--- ${label} XML Analysis ---`);

  const buffer = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml')!.async('string');

  // Find paragraphs containing '046762'
  const pos = docXml.indexOf('046762');
  if (pos === -1) {
    console.log('Target paragraph not found');
    return;
  }

  const beforeTerm = docXml.substring(0, pos);
  const paraStart = beforeTerm.lastIndexOf('<w:p ');
  const afterParaStart = docXml.substring(paraStart);
  const paraEnd = afterParaStart.indexOf('</w:p>') + 6;
  const para = afterParaStart.substring(0, paraEnd);

  // Check structure
  const hasBegin = para.includes('fldCharType="begin"');
  const hasSep = para.includes('fldCharType="separate"');
  const hasEnd = para.includes('fldCharType="end"');
  const insCount = (para.match(/<w:ins /g) || []).length;
  const delCount = (para.match(/<w:del /g) || []).length;

  console.log(`  Field: ${hasBegin ? 'BEGIN' : ''} ${hasSep ? 'SEP' : ''} ${hasEnd ? 'END' : ''}`);
  console.log(`  INS: ${insCount}, DEL: ${delCount}`);

  // Check if INS is after END
  const endPos = para.indexOf('fldCharType="end"');
  if (endPos > 0) {
    const afterEnd = para.substring(endPos);
    if (afterEnd.includes('<w:ins ')) {
      console.log('  *** BUG: INS is AFTER field END ***');
    } else {
      console.log('  OK: INS is before field END');
    }
  }
}

async function checkDocStructure(doc: any, label: string) {
  console.log(`\n--- ${label} In-Memory Structure ---`);

  const paragraphs = doc.getAllParagraphs();

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (!para) continue;

    const content = para.getContent();
    let hasHyperlinkField = false;
    let has046762 = false;

    // Check for HYPERLINK field with 046762
    for (const item of content) {
      if (!item) continue;
      const typeName = item.constructor.name;

      if (typeName === 'Run') {
        const run = item as any;
        const runContent = run.getContent();
        for (const c of runContent) {
          if (c.type === 'instructionText' && c.value && c.value.includes('HYPERLINK')) {
            hasHyperlinkField = true;
          }
          if (c.type === 'text' && c.value && c.value.includes('046762')) {
            has046762 = true;
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
              if ((c.type === 'text' || c.type === 'delText') && c.value && c.value.includes('046762')) {
                has046762 = true;
              }
            }
          }
        }
      }
    }

    if (!hasHyperlinkField || !has046762) continue;

    console.log(`  Paragraph ${i}: ${content.length} items`);

    // Show structure
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
    console.log(`  Structure: ${structure}`);
  }
}

async function main() {
  console.log('='.repeat(60));
  console.log('TEMPLATE_UI FULL PROCESSING SIMULATION');
  console.log('='.repeat(60));

  // Step 1: Analyze original document XML
  console.log('\n========================================');
  console.log('STEP 1: Original Document XML');
  console.log('========================================');
  await analyzeXml('Original_16.docx', 'Original');

  // Step 2: Load with revisionHandling: 'preserve' (like Template_UI)
  console.log('\n========================================');
  console.log('STEP 2: Load with preserve');
  console.log('========================================');
  const doc = await Document.load('Original_16.docx', {
    revisionHandling: 'preserve',
    strictParsing: false
  });
  await checkDocStructure(doc, 'After Load');

  // Step 3: Enable track changes (like Template_UI does)
  console.log('\n========================================');
  console.log('STEP 3: Enable track changes');
  console.log('========================================');
  doc.enableTrackChanges({
    author: 'Doc Hub',
    trackFormatting: true,
    showInsertionsAndDeletions: true,
  });
  await checkDocStructure(doc, 'After enableTrackChanges');

  // Step 4: Defragment hyperlinks
  console.log('\n========================================');
  console.log('STEP 4: Defragment hyperlinks');
  console.log('========================================');
  const merged = doc.defragmentHyperlinks({
    resetFormatting: true,
    cleanupRelationships: true,
  });
  console.log(`  Merged ${merged} fragmented hyperlinks`);
  await checkDocStructure(doc, 'After defragmentHyperlinks');

  // Step 5: Simulate hyperlink extraction (reads content)
  console.log('\n========================================');
  console.log('STEP 5: Extract hyperlinks');
  console.log('========================================');
  const paragraphs = doc.getAllParagraphs();
  for (const para of paragraphs) {
    const content = para.getContent();
    for (const item of content) {
      // Duck-type check for Hyperlink
      if (item && typeof (item as any).getUrl === 'function') {
        (item as any).getUrl();
        (item as any).getText();
      }
      // Check inside revisions
      if (item && typeof (item as any).getContent === 'function') {
        const revContent = (item as any).getContent();
        if (Array.isArray(revContent)) {
          for (const inner of revContent) {
            if (inner && typeof inner.getUrl === 'function') {
              inner.getUrl();
              inner.getText();
            }
          }
        }
      }
    }
  }
  await checkDocStructure(doc, 'After hyperlink extraction');

  // Step 6: Apply styles (calls clearDirectFormattingConflicts)
  console.log('\n========================================');
  console.log('STEP 6: Apply styles');
  console.log('========================================');
  doc.applyStyles({
    normal: {
      run: { font: 'Verdana', size: 12, color: '000000' }
    }
  });
  await checkDocStructure(doc, 'After applyStyles');

  // Step 7: Standardize hyperlink formatting
  console.log('\n========================================');
  console.log('STEP 7: Standardize hyperlink formatting');
  console.log('========================================');
  // Simulate what Template_UI does - iterate and format hyperlinks
  const allParas = doc.getAllParagraphs();
  for (const para of allParas) {
    const content = para.getContent();
    for (const item of content) {
      // Check for Hyperlink
      if (item && typeof (item as any).setFormatting === 'function' && typeof (item as any).getUrl === 'function') {
        (item as any).setFormatting({
          font: 'Verdana',
          size: 12,
          color: '0000FF',
          underline: 'single',
          bold: false,
          italic: false,
        }, { replace: true });
      }
    }
  }
  await checkDocStructure(doc, 'After hyperlink formatting');

  // Step 8: Save document
  console.log('\n========================================');
  console.log('STEP 8: Save document');
  console.log('========================================');
  await doc.save('test_full_simulation.docx');
  console.log('  Saved to test_full_simulation.docx');

  // Step 9: Analyze saved document XML
  console.log('\n========================================');
  console.log('STEP 9: Analyze saved document');
  console.log('========================================');
  await analyzeXml('test_full_simulation.docx', 'After Save');

  // Step 10: Reload and check structure
  console.log('\n========================================');
  console.log('STEP 10: Reload and verify');
  console.log('========================================');
  const doc2 = await Document.load('test_full_simulation.docx', {
    revisionHandling: 'preserve',
    strictParsing: false
  });
  await checkDocStructure(doc2, 'After Reload');
  await analyzeXml('test_full_simulation.docx', 'Final');

  doc.dispose();
  doc2.dispose();

  console.log('\n' + '='.repeat(60));
  console.log('SIMULATION COMPLETE');
  console.log('='.repeat(60));
}

main().catch(console.error);
