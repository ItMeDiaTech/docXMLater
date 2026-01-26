/**
 * Analyze the simulation output document
 */
import * as fs from 'fs';
import JSZip from 'jszip';

async function analyze() {
  // Analyze BOTH test files
  const files = ['test_full_simulation.docx', 'test_accept_revisions.docx', 'Processed_16.docx'];

  for (const file of files) {
    console.log(`\n=== Analyzing ${file} ===`);;

  if (!fs.existsSync('test_full_simulation.docx')) {
    console.log('File not found!');
    return;
  }

  const buffer = fs.readFileSync('test_full_simulation.docx');
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml')!.async('string');

  console.log('Document XML size:', docXml.length);

  const pos = docXml.indexOf('046762');
  console.log('Position of 046762:', pos);

  if (pos === -1) {
    console.log('Not found! Checking for Reconsideration...');
    const recon = docXml.indexOf('Reconsideration');
    console.log('Reconsideration position:', recon);

    // Check total revisions
    const insCount = (docXml.match(/<w:ins /g) || []).length;
    const delCount = (docXml.match(/<w:del /g) || []).length;
    console.log('Total INS:', insCount, 'Total DEL:', delCount);

    // Check for HYPERLINK
    const hyperlinkCount = (docXml.match(/HYPERLINK/g) || []).length;
    console.log('HYPERLINK occurrences:', hyperlinkCount);
    return;
  }

  // Find the enclosing paragraph by looking for </w:p> after the term
  // and <w:p from before
  const afterTerm = docXml.substring(pos);
  const paraEndFromTerm = afterTerm.indexOf('</w:p>');
  const actualEnd = pos + paraEndFromTerm + 6;

  // Now search backwards from the term position for <w:p
  const beforeTerm = docXml.substring(0, pos);
  let paraStart = beforeTerm.lastIndexOf('<w:p ');

  // Verify this is actually the start of our paragraph by checking
  // there's no </w:p> between paraStart and pos
  const between = docXml.substring(paraStart, pos);
  if (between.includes('</w:p>')) {
    // We found a different paragraph, need to look again
    // Find the last <w:p after the last </w:p> in between
    const lastParaEnd = between.lastIndexOf('</w:p>');
    const fromLastEnd = between.substring(lastParaEnd);
    const nextParaInBetween = fromLastEnd.indexOf('<w:p');
    if (nextParaInBetween !== -1) {
      paraStart = paraStart + lastParaEnd + nextParaInBetween;
    }
  }

  const para = docXml.substring(paraStart, actualEnd);

  console.log('\nParagraph containing 046762:');
  console.log('Length:', para.length);
  console.log('Has BEGIN:', para.includes('fldCharType="begin"'));
  console.log('Has SEP:', para.includes('fldCharType="separate"'));
  console.log('Has END:', para.includes('fldCharType="end"'));
  console.log('INS count:', (para.match(/<w:ins /g) || []).length);
  console.log('DEL count:', (para.match(/<w:del /g) || []).length);

  // Format for readability
  const formatted = para
    .replace(/></g, '>\n<')
    .split('\n')
    .map((line: string, i: number) => i.toString().padStart(3) + ': ' + line)
    .join('\n');

  console.log('\n--- Paragraph XML ---');
  console.log(formatted);
}

analyze().catch(console.error);
