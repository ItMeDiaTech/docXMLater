/**
 * TEST: What happens if we accept revisions then make modifications?
 * This might reproduce the Processed_16.docx pattern
 */

import { Document, Hyperlink, Revision, Run } from './src/index';
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
  console.log(`  INS count: ${elements.filter(e => e.type.includes('ins')).length}`);
  console.log(`  DEL count: ${elements.filter(e => e.type.includes('del')).length}`);
  console.log(`  Structure: ${elements.map(e => e.type).join(' ')}`);
}

async function main() {
  console.log('='.repeat(70));
  console.log('TEST: ACCEPT THEN MODIFY');
  console.log('='.repeat(70));

  // Test 1: Accept revisions then enable track changes and modify
  console.log('\n' + '='.repeat(50));
  console.log('SCENARIO 1: Accept revisions, then enable track changes');
  console.log('='.repeat(50));

  const doc1 = await Document.load('Original_16.docx', {
    revisionHandling: 'preserve',
    strictParsing: false
  });

  console.log('\n1. Loaded document');

  // Accept all revisions
  if (typeof (doc1 as any).acceptAllRevisions === 'function') {
    await (doc1 as any).acceptAllRevisions();
    console.log('2. Accepted all revisions');
  }

  // Enable track changes
  doc1.enableTrackChanges({
    author: 'Doc Hub',
    trackFormatting: true,
    showInsertionsAndDeletions: true,
  });
  console.log('3. Enabled track changes');

  // Apply styles (might create new revisions)
  doc1.applyStyles({
    normal: {
      run: { font: 'Verdana', size: 12, color: '000000' }
    }
  });
  console.log('4. Applied styles');

  // Save
  await doc1.save('test_accept_then_modify.docx');
  const buffer1 = fs.readFileSync('test_accept_then_modify.docx');
  await analyzeXmlStructure(buffer1, 'Scenario 1');

  doc1.dispose();

  // Test 2: Load with acceptRevisions option
  console.log('\n' + '='.repeat(50));
  console.log('SCENARIO 2: Load with acceptRevisions: true');
  console.log('='.repeat(50));

  const doc2 = await Document.load('Original_16.docx', {
    acceptRevisions: true,
    strictParsing: false
  });

  console.log('\n1. Loaded with acceptRevisions: true');

  // Enable track changes
  doc2.enableTrackChanges({
    author: 'Doc Hub',
    trackFormatting: true,
    showInsertionsAndDeletions: true,
  });
  console.log('2. Enabled track changes');

  // Apply styles
  doc2.applyStyles({
    normal: {
      run: { font: 'Verdana', size: 12, color: '000000' }
    }
  });
  console.log('3. Applied styles');

  // Save
  await doc2.save('test_load_accept_modify.docx');
  const buffer2 = fs.readFileSync('test_load_accept_modify.docx');
  await analyzeXmlStructure(buffer2, 'Scenario 2');

  doc2.dispose();

  // Compare with Processed_16.docx
  console.log('\n' + '='.repeat(50));
  console.log('COMPARISON');
  console.log('='.repeat(50));

  const processedBuffer = fs.readFileSync('Processed_16.docx');
  await analyzeXmlStructure(processedBuffer, 'Processed_16');

  console.log('\n' + '='.repeat(70));
  console.log('TEST COMPLETE');
  console.log('='.repeat(70));
}

main().catch(console.error);
