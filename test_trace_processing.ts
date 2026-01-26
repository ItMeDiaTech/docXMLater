/**
 * DETAILED DIAGNOSTIC: Trace paragraph content at every step
 * Find where the bug occurs
 */

import { Document, Hyperlink, Revision, Style } from './src/index';
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

  // Find enclosing paragraph
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
  console.log(`  Structure: ${elements.map(e => e.type).join(' ')}`);
}

function analyzeInMemoryStructure(doc: any, label: string): void {
  console.log(`\n--- ${label} (In-Memory) ---`);

  const paragraphs = doc.getAllParagraphs();

  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    if (!para) continue;

    const content = para.getContent();
    let has046762 = false;

    // Check if this paragraph contains 046762
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

    // Build structure string
    let structure = '';
    let insCount = 0;
    let delCount = 0;

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
        if (type === 'insert') insCount++;
        if (type === 'delete') delCount++;
      } else {
        structure += `<${typeName}>`;
      }
    }

    console.log(`  Content items: ${content.length}`);
    console.log(`  INS count: ${insCount}, DEL count: ${delCount}`);
    console.log(`  Structure: ${structure}`);
    break;
  }
}

async function main() {
  console.log('='.repeat(70));
  console.log('DETAILED PROCESSING TRACE');
  console.log('='.repeat(70));

  // 1. Analyze original file first
  console.log('\n' + '='.repeat(50));
  console.log('STEP 0: ORIGINAL FILE');
  console.log('='.repeat(50));
  const originalBuffer = fs.readFileSync('Original_16.docx');
  await analyzeXmlStructure(originalBuffer, 'Original');

  // 2. Load document
  console.log('\n' + '='.repeat(50));
  console.log('STEP 1: LOAD DOCUMENT');
  console.log('='.repeat(50));
  const doc = await Document.load('Original_16.docx', {
    revisionHandling: 'preserve',
    strictParsing: false
  });
  analyzeInMemoryStructure(doc, 'After Load');

  // 3. Enable track changes
  console.log('\n' + '='.repeat(50));
  console.log('STEP 2: ENABLE TRACK CHANGES');
  console.log('='.repeat(50));
  doc.enableTrackChanges({
    author: 'Doc Hub',
    trackFormatting: true,
    showInsertionsAndDeletions: true,
  });
  analyzeInMemoryStructure(doc, 'After enableTrackChanges');

  // 4. Defragment hyperlinks
  console.log('\n' + '='.repeat(50));
  console.log('STEP 3: DEFRAGMENT HYPERLINKS');
  console.log('='.repeat(50));
  const merged = doc.defragmentHyperlinks({
    resetFormatting: true,
    cleanupRelationships: true,
  });
  console.log(`  Merged ${merged} hyperlinks`);
  analyzeInMemoryStructure(doc, 'After defragmentHyperlinks');

  // 5. Standardize hyperlink formatting (simulated)
  console.log('\n' + '='.repeat(50));
  console.log('STEP 4: STANDARDIZE HYPERLINK FORMATTING');
  console.log('='.repeat(50));
  const paragraphs = doc.getAllParagraphs();
  let standardizedCount = 0;
  for (const para of paragraphs) {
    if (!para) continue;
    const content = para.getContent();
    for (const item of content) {
      if (item instanceof Hyperlink) {
        item.setFormatting({
          font: 'Verdana',
          size: 12,
          color: '0000FF',
          underline: 'single',
          bold: false,
          italic: false,
        }, { replace: true });
        standardizedCount++;
      } else if (item instanceof Revision) {
        const revContent = item.getContent();
        for (const inner of revContent) {
          if (inner instanceof Hyperlink) {
            inner.setFormatting({
              font: 'Verdana',
              size: 12,
              color: '0000FF',
              underline: 'single',
              bold: false,
              italic: false,
            }, { replace: true });
            standardizedCount++;
          }
        }
      }
    }
  }
  console.log(`  Standardized ${standardizedCount} hyperlinks`);
  analyzeInMemoryStructure(doc, 'After standardizeHyperlinkFormatting');

  // 6. Apply styles
  console.log('\n' + '='.repeat(50));
  console.log('STEP 5: APPLY STYLES');
  console.log('='.repeat(50));
  doc.applyStyles({
    normal: {
      run: { font: 'Verdana', size: 12, color: '000000' }
    }
  });
  analyzeInMemoryStructure(doc, 'After applyStyles');

  // 7. Save to buffer
  console.log('\n' + '='.repeat(50));
  console.log('STEP 6: SAVE TO BUFFER');
  console.log('='.repeat(50));
  const buffer = await doc.toBuffer();
  await analyzeXmlStructure(buffer, 'After toBuffer');

  // 8. Save to file
  console.log('\n' + '='.repeat(50));
  console.log('STEP 7: SAVE TO FILE');
  console.log('='.repeat(50));
  await doc.save('test_trace_output.docx');
  const savedBuffer = fs.readFileSync('test_trace_output.docx');
  await analyzeXmlStructure(savedBuffer, 'After save');

  // 9. Compare with Processed_16.docx
  console.log('\n' + '='.repeat(50));
  console.log('COMPARISON: PROCESSED_16.DOCX');
  console.log('='.repeat(50));
  const processedBuffer = fs.readFileSync('Processed_16.docx');
  await analyzeXmlStructure(processedBuffer, 'Processed_16');

  doc.dispose();

  console.log('\n' + '='.repeat(70));
  console.log('TRACE COMPLETE');
  console.log('='.repeat(70));
}

main().catch(console.error);
