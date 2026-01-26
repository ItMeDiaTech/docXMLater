/**
 * Test if acceptAllRevisions causes the field code corruption
 */

import { Document } from './src/index';
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
      console.log('  OK: INS is before field END (or no INS after END)');
    }
  }

  // Show element order
  console.log('\n  Element positions in paragraph:');
  const elements: { pos: number; type: string }[] = [];

  let idx = 0;
  while ((idx = para.indexOf('fldCharType=', idx)) !== -1) {
    const typeStart = para.indexOf('"', idx) + 1;
    const typeEnd = para.indexOf('"', typeStart);
    const fldType = para.substring(typeStart, typeEnd);
    elements.push({ pos: idx, type: `fldChar:${fldType}` });
    idx++;
  }

  idx = 0;
  while ((idx = para.indexOf('<w:ins ', idx)) !== -1) {
    elements.push({ pos: idx, type: 'w:ins' });
    idx++;
  }

  idx = 0;
  while ((idx = para.indexOf('<w:del ', idx)) !== -1) {
    elements.push({ pos: idx, type: 'w:del' });
    idx++;
  }

  elements.sort((a, b) => a.pos - b.pos);
  elements.forEach(e => console.log(`    ${e.pos}: ${e.type}`));
}

async function checkDocStructure(doc: any, label: string) {
  console.log(`\n--- ${label} In-Memory Structure ---`);

  const paragraphs = doc.getAllParagraphs();

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (!para) continue;

    const content = para.getContent();
    let has046762 = false;

    // Check for 046762
    for (const item of content) {
      if (!item) continue;
      const typeName = item.constructor.name;

      if (typeName === 'Run') {
        const run = item as any;
        const runContent = run.getContent();
        for (const c of runContent) {
          if ((c.type === 'text' || c.type === 'delText') && c.value && c.value.includes('046762')) {
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
              if ((c.type === 'text' || c.type === 'delText') && c.value && c.value.includes('046762')) {
                has046762 = true;
              }
            }
          }
        }
      }
    }

    if (!has046762) continue;

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
    break;
  }
}

async function main() {
  console.log('='.repeat(60));
  console.log('TEST: ACCEPT ALL REVISIONS EFFECT ON FIELD CODES');
  console.log('='.repeat(60));

  // Step 1: Load with revisions preserved
  console.log('\n========================================');
  console.log('STEP 1: Load with preserve');
  console.log('========================================');
  const doc = await Document.load('Original_16.docx', {
    revisionHandling: 'preserve',
    strictParsing: false
  });
  await checkDocStructure(doc, 'After Load');

  // Step 2: Accept all revisions
  console.log('\n========================================');
  console.log('STEP 2: Accept all revisions');
  console.log('========================================');
  if (typeof (doc as any).acceptAllRevisions === 'function') {
    await (doc as any).acceptAllRevisions();
    console.log('  acceptAllRevisions() called');
  } else {
    console.log('  acceptAllRevisions not available!');
  }
  await checkDocStructure(doc, 'After acceptAllRevisions');

  // Step 3: Save
  console.log('\n========================================');
  console.log('STEP 3: Save document');
  console.log('========================================');
  await doc.save('test_accept_revisions.docx');
  console.log('  Saved to test_accept_revisions.docx');

  // Step 4: Analyze saved document
  console.log('\n========================================');
  console.log('STEP 4: Analyze saved document');
  console.log('========================================');
  await analyzeXml('test_accept_revisions.docx', 'After acceptAllRevisions + Save');

  doc.dispose();

  console.log('\n' + '='.repeat(60));
  console.log('TEST COMPLETE');
  console.log('='.repeat(60));
}

main().catch(console.error);
