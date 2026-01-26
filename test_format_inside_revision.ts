/**
 * TEST: What happens when we apply formatting to runs inside revisions with track changes?
 * This might reproduce the Processed_16.docx bug
 */

import { Document, Revision, Run } from './src/index';
import * as fs from 'fs';
import JSZip from 'jszip';

interface ElementInfo {
  pos: number;
  type: string;
}

async function analyzeXmlStructure(buffer: Buffer | ArrayBuffer, label: string): Promise<void> {
  console.log(`\n--- ${label} (XML) ---`);

  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml')!.async('string');

  const pos = docXml.indexOf('046762');
  if (pos === -1) {
    console.log('  046762 not found in XML');
    return;
  }

  const afterTerm = docXml.substring(pos);
  const paraEndFromTerm = afterTerm.indexOf('</w:p>');
  const actualEnd = pos + paraEndFromTerm + 6;

  const beforeTerm = docXml.substring(0, pos);
  let paraStart = beforeTerm.lastIndexOf('<w:p ');

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

  const elements: ElementInfo[] = [];

  let idx = 0;
  while ((idx = para.indexOf('fldCharType=', idx)) !== -1) {
    const typeStart = para.indexOf('"', idx) + 1;
    const typeEnd = para.indexOf('"', typeStart);
    const fldType = para.substring(typeStart, typeEnd);
    elements.push({ pos: idx, type: `[${fldType}]` });
    idx++;
  }

  idx = 0;
  while ((idx = para.indexOf('<w:ins ', idx)) !== -1) {
    const idMatch = para.substring(idx).match(/w:id="(\d+)"/);
    const id = idMatch ? idMatch[1] : '?';
    elements.push({ pos: idx, type: `{ins id=${id}}` });
    idx++;
  }

  idx = 0;
  while ((idx = para.indexOf('<w:del ', idx)) !== -1) {
    const idMatch = para.substring(idx).match(/w:id="(\d+)"/);
    const id = idMatch ? idMatch[1] : '?';
    elements.push({ pos: idx, type: `{del id=${id}}` });
    idx++;
  }

  elements.sort((a, b) => a.pos - b.pos);
  console.log(`  INS count: ${elements.filter(e => e.type.includes('ins')).length}`);
  console.log(`  DEL count: ${elements.filter(e => e.type.includes('del')).length}`);
  console.log(`  Structure: ${elements.map(e => e.type).join(' ')}`);
}

function analyzeInMemory(doc: any, label: string): void {
  console.log(`\n--- ${label} (In-Memory) ---`);

  const paragraphs = doc.getAllParagraphs();

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (!para) continue;

    const content = para.getContent();
    let has046762 = false;

    // Check if paragraph contains 046762
    for (const item of content) {
      if (!item) continue;

      if (item instanceof Revision) {
        const revContent = item.getContent();
        for (const child of revContent) {
          if (child && typeof (child as any).getContent === 'function') {
            const childContent = (child as any).getContent();
            for (const c of childContent) {
              if ((c.type === 'text' || c.type === 'delText') && c.value?.includes('046762')) {
                has046762 = true;
              }
            }
          }
        }
      } else if (typeof (item as any).getContent === 'function') {
        const runContent = (item as any).getContent();
        if (Array.isArray(runContent)) {
          for (const c of runContent) {
            if ((c.type === 'text' || c.type === 'delText') && c.value?.includes('046762')) {
              has046762 = true;
            }
          }
        }
      }
    }

    if (!has046762) continue;

    // Build structure
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
      } else if (item instanceof Revision) {
        const rev = item;
        const type = rev.getType();
        const id = rev.getId();
        structure += `{${type} id=${id}}`;
      } else {
        structure += `<${typeName}>`;
      }
    }

    console.log(`  Content items: ${content.length}`);
    console.log(`  Structure: ${structure}`);
    break;
  }
}

/**
 * Get all runs from paragraph including those inside revisions (like Template_UI does)
 */
function getAllRunsFromParagraph(para: any): Run[] {
  const allRuns: Run[] = [];
  const content = para.getContent();

  for (const item of content) {
    if (item instanceof Run) {
      allRuns.push(item);
    } else if (item instanceof Revision) {
      const revRuns = item.getRuns();
      allRuns.push(...revRuns);
    }
  }

  return allRuns;
}

async function main() {
  console.log('='.repeat(70));
  console.log('TEST: FORMAT INSIDE REVISION');
  console.log('='.repeat(70));

  // Original for reference
  const originalBuffer = fs.readFileSync('Original_16.docx');
  await analyzeXmlStructure(originalBuffer, 'Original');

  // Scenario 1: Load, enable track changes, apply formatting to all runs
  console.log('\n' + '='.repeat(50));
  console.log('SCENARIO 1: Apply formatting to runs inside revisions');
  console.log('(This simulates what Template_UI does with getAllRunsFromParagraph)');
  console.log('='.repeat(50));

  const doc1 = await Document.load('Original_16.docx', {
    revisionHandling: 'preserve',
    strictParsing: false
  });

  console.log('\n1. Loaded document');
  analyzeInMemory(doc1, 'After Load');

  // Enable track changes
  doc1.enableTrackChanges({
    author: 'Doc Hub',
    trackFormatting: true,
    showInsertionsAndDeletions: true,
  });
  console.log('2. Enabled track changes');

  // Apply formatting to all runs including inside revisions (like Template_UI)
  console.log('3. Applying formatting to all runs (including inside revisions)...');
  const paragraphs = doc1.getAllParagraphs();
  let modifiedCount = 0;

  for (const para of paragraphs) {
    if (!para) continue;

    const runs = getAllRunsFromParagraph(para);
    for (const run of runs) {
      // Skip hyperlink-styled runs (like Template_UI)
      if (run.isHyperlinkStyled()) {
        continue;
      }

      // Apply formatting (like Template_UI's applyCustomFormatting)
      run.setFont('Verdana');
      run.setSize(12);
      run.setColor('000000');
      run.setUnderline(false);
      run.setBold(false);
      run.setItalic(false);
      modifiedCount++;
    }
  }
  console.log(`   Modified ${modifiedCount} runs`);

  analyzeInMemory(doc1, 'After Formatting');

  // Save
  await doc1.save('test_format_inside_revision.docx');
  const buffer1 = fs.readFileSync('test_format_inside_revision.docx');
  await analyzeXmlStructure(buffer1, 'Scenario 1 Output');

  doc1.dispose();

  // Compare with Processed_16.docx
  console.log('\n' + '='.repeat(50));
  console.log('COMPARISON: PROCESSED_16.DOCX');
  console.log('='.repeat(50));

  const processedBuffer = fs.readFileSync('Processed_16.docx');
  await analyzeXmlStructure(processedBuffer, 'Processed_16');

  console.log('\n' + '='.repeat(70));
  console.log('TEST COMPLETE');
  console.log('='.repeat(70));
}

main().catch(console.error);
