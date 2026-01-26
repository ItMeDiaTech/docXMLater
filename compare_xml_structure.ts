/**
 * Compare XML structure between Original and Processed documents
 */

import * as fs from 'fs';
import JSZip from 'jszip';

async function extractParagraphs(filePath: string, label: string) {
  console.log('');
  console.log('='.repeat(60));
  console.log(label);
  console.log('='.repeat(60));

  const buffer = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml')!.async('string');

  console.log('Document.xml size:', docXml.length, 'chars');

  // Find paragraphs containing HYPERLINK with revisions
  const paras = docXml.split('<w:p ');
  console.log('Total paragraph splits:', paras.length);

  let foundCount = 0;
  for (let i = 0; i < paras.length; i++) {
    const para = paras[i];
    if (!para) continue;
    // Look for HYPERLINK field codes with ins/del revisions
    if (para.includes('HYPERLINK') && (para.includes('<w:ins ') || para.includes('<w:del '))) {
      foundCount++;
      if (foundCount > 3) continue; // Limit output

      console.log('');
      console.log('--- Paragraph ' + i + ': HYPERLINK field with revisions ---');

      // Check structure markers
      const hasBegin = para.includes('fldCharType="begin"');
      const hasSep = para.includes('fldCharType="separate"');
      const hasEnd = para.includes('fldCharType="end"');
      const insCount = (para.match(/<w:ins /g) || []).length;
      const delCount = (para.match(/<w:del /g) || []).length;

      console.log('Field chars: ' + (hasBegin ? 'BEGIN ' : '') + (hasSep ? 'SEP ' : '') + (hasEnd ? 'END' : ''));
      console.log('Revisions: INS=' + insCount + ', DEL=' + delCount);

      // Find positions
      const beginPos = para.indexOf('fldCharType="begin"');
      const sepPos = para.indexOf('fldCharType="separate"');
      const endPos = para.indexOf('fldCharType="end"');

      console.log('Positions: BEGIN=' + beginPos + ', SEP=' + sepPos + ', END=' + endPos);

      // Check what comes after END
      if (endPos > 0) {
        const afterEnd = para.substring(endPos + 20);
        const hasInsAfterEnd = afterEnd.includes('<w:ins ');
        const hasDelAfterEnd = afterEnd.includes('<w:del ');
        console.log('After END: INS=' + hasInsAfterEnd + ', DEL=' + hasDelAfterEnd);
        if (hasInsAfterEnd) {
          console.log('*** BUG: INS revision appears AFTER field end! ***');
        }
      }
    }
  }
  console.log('');
  console.log('Total HYPERLINK paragraphs with revisions: ' + foundCount);
}

async function findSpecificParagraphs(filePath: string, label: string) {
  console.log('');
  console.log('='.repeat(60));
  console.log(label + ' - Looking for IBR and Reconsideration hyperlinks');
  console.log('='.repeat(60));

  const buffer = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml')!.async('string');

  // Look for paragraphs with these specific hyperlinks
  const searchTerms = ['IBR', 'Reconsideration', '075115', '046762'];

  for (const term of searchTerms) {
    const pos = docXml.indexOf(term);
    if (pos !== -1) {
      // Find the paragraph containing this term
      const beforeTerm = docXml.substring(0, pos);
      const paraStart = beforeTerm.lastIndexOf('<w:p ');
      const afterParaStart = docXml.substring(paraStart);
      const paraEnd = afterParaStart.indexOf('</w:p>') + 6;
      const para = afterParaStart.substring(0, paraEnd);

      console.log('');
      console.log('Found: ' + term);
      console.log('Paragraph length:', para.length, 'chars');
      console.log('Has BEGIN:', para.includes('fldCharType="begin"'));
      console.log('Has SEP:', para.includes('fldCharType="separate"'));
      console.log('Has END:', para.includes('fldCharType="end"'));
      console.log('INS count:', (para.match(/<w:ins /g) || []).length);
      console.log('DEL count:', (para.match(/<w:del /g) || []).length);
      console.log('delText count:', (para.match(/<w:delText/g) || []).length);
      console.log('delInstrText count:', (para.match(/<w:delInstrText/g) || []).length);

      // Check if HYPERLINK is in this paragraph
      if (para.includes('HYPERLINK')) {
        console.log('Contains HYPERLINK field code: YES');

        // Show the actual structure elements in order
        console.log('');
        console.log('Structure elements in order:');

        // Simple order check - extract key elements and their positions
        interface StructElement {
          pos: number;
          type: string;
          content?: string;
        }
        const elements: StructElement[] = [];

        // Find fldChar elements
        let idx = 0;
        while ((idx = para.indexOf('fldCharType=', idx)) !== -1) {
          const typeStart = para.indexOf('"', idx) + 1;
          const typeEnd = para.indexOf('"', typeStart);
          const fldType = para.substring(typeStart, typeEnd);
          elements.push({ pos: idx, type: 'fldChar:' + fldType });
          idx++;
        }

        // Find ins/del elements
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

        // Sort by position and print
        elements.sort((a, b) => a.pos - b.pos);
        elements.forEach(e => console.log('  ' + e.pos + ': ' + e.type));
      } else {
        console.log('Contains HYPERLINK field code: NO');
      }
    } else {
      console.log('');
      console.log('NOT FOUND in document: ' + term);
    }
  }
}

async function dumpParagraphXml(filePath: string, searchTerm: string) {
  console.log('');
  console.log('='.repeat(60));
  console.log('Dumping paragraph XML for: ' + searchTerm);
  console.log('File: ' + filePath);
  console.log('='.repeat(60));

  const buffer = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file('word/document.xml')!.async('string');

  const pos = docXml.indexOf(searchTerm);
  if (pos !== -1) {
    const beforeTerm = docXml.substring(0, pos);
    const paraStart = beforeTerm.lastIndexOf('<w:p ');
    const afterParaStart = docXml.substring(paraStart);
    const paraEnd = afterParaStart.indexOf('</w:p>') + 6;
    const para = afterParaStart.substring(0, paraEnd);

    // Format the XML for readability
    const formatted = para
      .replace(/></g, '>\n<')
      .split('\n')
      .map((line, i) => `${i.toString().padStart(3)}: ${line}`)
      .join('\n');

    console.log(formatted);
  } else {
    console.log('NOT FOUND');
  }
}

async function main() {
  try {
    await extractParagraphs('Original_16.docx', 'ORIGINAL');
  } catch (e) {
    console.log('Error with Original_16.docx:', e);
  }

  try {
    await extractParagraphs('Processed_16.docx', 'PROCESSED');
  } catch (e) {
    console.log('Error with Processed_16.docx:', e);
  }

  // Now look for specific paragraphs
  try {
    await findSpecificParagraphs('Original_16.docx', 'ORIGINAL');
  } catch (e) {
    console.log('Error:', e);
  }

  try {
    await findSpecificParagraphs('Processed_16.docx', 'PROCESSED');
  } catch (e) {
    console.log('Error:', e);
  }

  // Dump the specific paragraph XML
  console.log('\n\n========== DETAILED XML DUMP ==========\n');
  await dumpParagraphXml('Original_16.docx', '046762');
  await dumpParagraphXml('Processed_16.docx', '046762');
}

main().catch(console.error);
