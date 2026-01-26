/**
 * Analyze multiple docx files to compare their 046762 paragraph structure
 */
import * as fs from 'fs';
import JSZip from 'jszip';

async function analyzeFile(file: string) {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`FILE: ${file}`);
  console.log('='.repeat(60));

  if (!fs.existsSync(file)) {
    console.log('  File not found!');
    return;
  }

  const buffer = fs.readFileSync(file);
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml')!.async('string');

  console.log(`  Document XML size: ${docXml.length}`);

  const pos = docXml.indexOf('046762');
  console.log(`  Position of 046762: ${pos}`);

  if (pos === -1) {
    console.log('  046762 not found in document!');
    return;
  }

  // Find the enclosing paragraph
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

  console.log(`\n  Paragraph structure:`);
  console.log(`  Has BEGIN: ${para.includes('fldCharType="begin"')}`);
  console.log(`  Has SEP: ${para.includes('fldCharType="separate"')}`);
  console.log(`  Has END: ${para.includes('fldCharType="end"')}`);
  console.log(`  INS count: ${(para.match(/<w:ins /g) || []).length}`);
  console.log(`  DEL count: ${(para.match(/<w:del /g) || []).length}`);

  // Check positions
  const beginPos = para.indexOf('fldCharType="begin"');
  const sepPos = para.indexOf('fldCharType="separate"');
  const endPos = para.indexOf('fldCharType="end"');

  console.log(`\n  Positions:`);
  console.log(`    BEGIN: ${beginPos}`);
  console.log(`    SEP: ${sepPos}`);
  console.log(`    END: ${endPos}`);

  // Find all INS positions
  const insPositions: number[] = [];
  let idx = 0;
  while ((idx = para.indexOf('<w:ins ', idx)) !== -1) {
    insPositions.push(idx);
    idx++;
  }

  if (insPositions.length > 0) {
    console.log(`    INS positions: ${insPositions.join(', ')}`);

    // Check if any INS is after END
    for (const insPos of insPositions) {
      if (insPos > endPos && endPos > 0) {
        console.log(`\n  *** BUG: INS at ${insPos} is AFTER END at ${endPos} ***`);
      }
    }
  }

  // Show element order
  console.log(`\n  Element order:`);
  const elements: { pos: number; type: string }[] = [];

  idx = 0;
  while ((idx = para.indexOf('fldCharType=', idx)) !== -1) {
    const typeStart = para.indexOf('"', idx) + 1;
    const typeEnd = para.indexOf('"', typeStart);
    const fldType = para.substring(typeStart, typeEnd);
    elements.push({ pos: idx, type: `[${fldType}]` });
    idx++;
  }

  idx = 0;
  while ((idx = para.indexOf('<w:ins ', idx)) !== -1) {
    // Find revision ID
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
  const structure = elements.map(e => e.type).join(' ');
  console.log(`    ${structure}`);
}

async function main() {
  const files = [
    'Original_16.docx',
    'test_full_simulation.docx',
    'test_exact_templateui.docx',
    'test_accept_revisions.docx',
    'Processed_16.docx'
  ];

  for (const file of files) {
    await analyzeFile(file);
  }

  console.log('\n' + '='.repeat(60));
  console.log('ANALYSIS COMPLETE');
  console.log('='.repeat(60));
}

main().catch(console.error);
