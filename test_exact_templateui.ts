/**
 * Exact Template_UI processing simulation
 * This should reproduce the exact processing steps Template_UI uses
 */

import { Document, Hyperlink, Revision, Style } from './src/index';
import * as fs from 'fs';
import JSZip from 'jszip';

async function analyzeXml(filePath: string, label: string) {
  console.log(`\n--- ${label} XML Analysis ---`);

  const buffer = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml')!.async('string');

  // Find paragraph containing '046762'
  const pos = docXml.indexOf('046762');
  if (pos === -1) {
    console.log('046762 not found!');
    return;
  }

  const afterTerm = docXml.substring(pos);
  const paraEndFromTerm = afterTerm.indexOf('</w:p>');
  const actualEnd = pos + paraEndFromTerm + 6;

  const beforeTerm = docXml.substring(0, pos);
  let paraStart = beforeTerm.lastIndexOf('<w:p ');

  // Verify correct paragraph
  const between = docXml.substring(paraStart, pos);
  if (between.includes('</w:p>')) {
    const lastParaEnd = between.lastIndexOf('</w:p>');
    const fromLastEnd = between.substring(lastParaEnd);
    const nextParaInBetween = fromLastEnd.indexOf('<w:p');
    if (nextParaInBetween !== -1) {
      paraStart = paraStart + lastParaEnd + nextParaInBetween;
    }
  }

  const para = docXml.substring(paraStart, actualEnd);

  console.log('  Has BEGIN:', para.includes('fldCharType="begin"'));
  console.log('  Has SEP:', para.includes('fldCharType="separate"'));
  console.log('  Has END:', para.includes('fldCharType="end"'));
  console.log('  INS count:', (para.match(/<w:ins /g) || []).length);
  console.log('  DEL count:', (para.match(/<w:del /g) || []).length);

  // Check position of INS relative to END
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
    break; // Only show first matching paragraph
  }
}

/**
 * Apply standard hyperlink formatting (mimics Template_UI's standardizeHyperlinkFormatting)
 */
function standardizeHyperlinkFormatting(doc: any) {
  let count = 0;
  const paragraphs = doc.getAllParagraphs();

  for (const para of paragraphs) {
    const content = para.getContent();

    for (const item of content) {
      // Case 1: Direct Hyperlink instances
      if (item instanceof Hyperlink) {
        item.setFormatting({
          font: 'Verdana',
          size: 12,
          color: '0000FF',
          underline: 'single',
          bold: false,
          italic: false,
        }, { replace: true });
        count++;
      }
      // Case 2: Hyperlinks inside Revision elements
      else if (item instanceof Revision) {
        const revisionContent = item.getContent();
        for (const revContent of revisionContent) {
          if (revContent instanceof Hyperlink) {
            revContent.setFormatting({
              font: 'Verdana',
              size: 12,
              color: '0000FF',
              underline: 'single',
              bold: false,
              italic: false,
            }, { replace: true });
            count++;
          }
        }
      }
    }
  }

  return count;
}

/**
 * Check if paragraph has complex field content
 */
function hasComplexFieldContent(para: any): boolean {
  try {
    for (const run of para.getRuns()) {
      const content = run.getContent();
      if (content.some((c: { type: string }) =>
        c.type === 'instructionText' || c.type === 'fieldChar'
      )) {
        return true;
      }
    }
    return false;
  } catch {
    return false;
  }
}

/**
 * Simulate removeExtraWhitespace (this modifies runs in paragraphs)
 */
function removeExtraWhitespace(doc: any) {
  const paragraphs = doc.getAllParagraphs();

  for (const para of paragraphs) {
    // Skip paragraphs with complex field content
    if (hasComplexFieldContent(para)) {
      continue;
    }

    const runs = para.getRuns();
    for (let i = 0; i < runs.length; i++) {
      const run = runs[i];
      const text = run.getText();
      if (!text) continue;

      // Collapse multiple spaces
      let cleaned = text.replace(/\s+/g, ' ');

      if (cleaned !== text) {
        run.setText(cleaned);
      }
    }
  }
}

async function main() {
  console.log('='.repeat(60));
  console.log('EXACT TEMPLATE_UI PROCESSING SIMULATION');
  console.log('='.repeat(60));

  // Step 1: Original analysis
  console.log('\n========================================');
  console.log('STEP 1: Original Document');
  console.log('========================================');
  await analyzeXml('Original_16.docx', 'Original');

  // Step 2: Load with revisionHandling: 'preserve'
  console.log('\n========================================');
  console.log('STEP 2: Load with preserve');
  console.log('========================================');
  const doc = await Document.load('Original_16.docx', {
    revisionHandling: 'preserve',
    strictParsing: false
  });
  await checkDocStructure(doc, 'After Load');

  // Step 3: Enable track changes (Template_UI does this)
  console.log('\n========================================');
  console.log('STEP 3: Enable track changes');
  console.log('========================================');
  doc.enableTrackChanges({
    author: 'Doc Hub',
    trackFormatting: true,
    showInsertionsAndDeletions: true,
  });
  await checkDocStructure(doc, 'After enableTrackChanges');

  // Step 4: Defragment hyperlinks (Template_UI does this)
  console.log('\n========================================');
  console.log('STEP 4: Defragment hyperlinks');
  console.log('========================================');
  const merged = doc.defragmentHyperlinks({
    resetFormatting: true,
    cleanupRelationships: true,
  });
  console.log(`  Merged ${merged} fragmented hyperlinks`);
  await checkDocStructure(doc, 'After defragmentHyperlinks');

  // Step 5: Standardize hyperlink formatting (Template_UI ALWAYS does this)
  console.log('\n========================================');
  console.log('STEP 5: Standardize hyperlink formatting');
  console.log('========================================');
  const standardized = standardizeHyperlinkFormatting(doc);
  console.log(`  Standardized ${standardized} hyperlinks`);
  await checkDocStructure(doc, 'After standardizeHyperlinkFormatting');

  // Step 6: Apply styles (Template_UI does this)
  console.log('\n========================================');
  console.log('STEP 6: Apply styles');
  console.log('========================================');
  doc.applyStyles({
    normal: {
      run: { font: 'Verdana', size: 12, color: '000000' }
    }
  });
  await checkDocStructure(doc, 'After applyStyles');

  // Step 7: Remove extra whitespace (Template_UI does this)
  console.log('\n========================================');
  console.log('STEP 7: Remove extra whitespace');
  console.log('========================================');
  removeExtraWhitespace(doc);
  await checkDocStructure(doc, 'After removeExtraWhitespace');

  // Step 8: Add Hyperlink style definition (Template_UI does this)
  console.log('\n========================================');
  console.log('STEP 8: Update Hyperlink style definition');
  console.log('========================================');
  const hyperlinkStyle = Style.create({
    styleId: 'Hyperlink',
    name: 'Hyperlink',
    type: 'character',
    runFormatting: {
      font: 'Verdana',
      size: 12,
      color: '0000FF',
      underline: 'single',
      bold: false,
      italic: false,
    },
  });
  doc.addStyle(hyperlinkStyle);
  await checkDocStructure(doc, 'After Hyperlink style update');

  // Step 9: Save document
  console.log('\n========================================');
  console.log('STEP 9: Save document');
  console.log('========================================');
  await doc.save('test_exact_templateui.docx');
  console.log('  Saved to test_exact_templateui.docx');

  // Step 10: Analyze saved document
  console.log('\n========================================');
  console.log('STEP 10: Analyze saved document');
  console.log('========================================');
  await analyzeXml('test_exact_templateui.docx', 'After Save');

  // Step 11: Reload and verify
  console.log('\n========================================');
  console.log('STEP 11: Reload and verify');
  console.log('========================================');
  const doc2 = await Document.load('test_exact_templateui.docx', {
    revisionHandling: 'preserve',
    strictParsing: false
  });
  await checkDocStructure(doc2, 'After Reload');

  doc.dispose();
  doc2.dispose();

  console.log('\n' + '='.repeat(60));
  console.log('SIMULATION COMPLETE');
  console.log('='.repeat(60));
}

main().catch(console.error);
