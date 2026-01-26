/**
 * Dump the exact XML structure of the 046762 paragraph from each file
 */

import * as fs from 'fs';
import JSZip from 'jszip';

async function dumpParagraph(filePath: string, label: string): Promise<void> {
  console.log('\n' + '='.repeat(70));
  console.log(`${label}: ${filePath}`);
  console.log('='.repeat(70));

  if (!fs.existsSync(filePath)) {
    console.log('  File not found');
    return;
  }

  const buffer = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml')!.async('string');

  const pos = docXml.indexOf('046762');
  if (pos === -1) {
    console.log('  046762 not found in document');
    return;
  }

  // Find enclosing paragraph
  const afterTerm = docXml.substring(pos);
  const paraEndFromTerm = afterTerm.indexOf('</w:p>');
  const actualEnd = pos + paraEndFromTerm + 6;

  const beforeTerm = docXml.substring(0, pos);
  let paraStart = beforeTerm.lastIndexOf('<w:p ');

  // Verify we have the correct paragraph
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

  // Format the XML with line numbers
  const formatted = para
    .replace(/></g, '>\n<')
    .split('\n')
    .map((line, i) => {
      // Highlight key elements
      let marker = '  ';
      if (line.includes('fldChar')) marker = '🔹';
      if (line.includes('w:ins ')) marker = '✅';
      if (line.includes('</w:ins>')) marker = '✅';
      if (line.includes('w:del ')) marker = '❌';
      if (line.includes('</w:del>')) marker = '❌';
      if (line.includes('046762')) marker = '🎯';
      return `${marker} ${(i + 1).toString().padStart(3)}: ${line.substring(0, 120)}${line.length > 120 ? '...' : ''}`;
    })
    .join('\n');

  console.log(formatted);
}

async function main() {
  // Dump original
  await dumpParagraph('Original_16.docx', 'ORIGINAL');

  // Dump my simulation output
  await dumpParagraph('test_trace_output.docx', 'MY SIMULATION');

  // Dump processed
  await dumpParagraph('Processed_16.docx', 'PROCESSED (BUG)');
}

main().catch(console.error);
