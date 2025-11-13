/**
 * Example: Ensure Blank Lines After 1x1 Tables
 *
 * This example demonstrates how to use the ensureBlankLinesAfter1x1Tables() method
 * to add preserved blank lines after all single-cell tables in a document.
 *
 * This is particularly useful when:
 * 1. You want consistent spacing after Header 2 tables
 * 2. You're processing documents with Template_UI and want to preserve spacing
 * 3. You need to ensure blank lines won't be removed by "remove blank lines" operations
 */

import { Document } from '../src';

async function main() {
  console.log('Loading document...');
  const doc = await Document.load('input.docx');

  console.log('\nBefore processing:');
  const tables = doc.getAllTables();
  console.log(`Total tables: ${tables.length}`);
  const oneByone = tables.filter(t => t.getRowCount() === 1 && t.getColumnCount() === 1);
  console.log(`1x1 tables: ${oneByone.length}`);

  // Ensure blank lines after all 1x1 tables
  console.log('\nProcessing 1x1 tables...');
  const result = doc.ensureBlankLinesAfter1x1Tables({
    spacingAfter: 120,       // 6pt spacing (default)
    markAsPreserved: true,   // Mark as preserved (default)
  });

  console.log('\nResults:');
  console.log(`Tables processed: ${result.tablesProcessed}`);
  console.log(`Blank lines added: ${result.blankLinesAdded}`);
  console.log(`Existing blank lines marked: ${result.existingLinesMarked}`);

  // Save the document
  console.log('\nSaving document...');
  await doc.save('output.docx');
  console.log('Done! Document saved to output.docx');

  // Cleanup
  doc.dispose();
}

// Example 2: Only process Header 2 tables
async function processHeader2TablesOnly() {
  const doc = await Document.load('input.docx');

  const result = doc.ensureBlankLinesAfter1x1Tables({
    filter: (table, index) => {
      const cell = table.getCell(0, 0);
      if (!cell) return false;

      // Check if cell contains a Header 2 paragraph
      return cell.getParagraphs().some(p => {
        const style = p.getStyle();
        return style === 'Heading2' || style === 'Heading 2' || style === 'Header2';
      });
    }
  });

  console.log(`Processed ${result.tablesProcessed} Header 2 tables`);
  await doc.save('output.docx');
  doc.dispose();
}

// Example 3: Custom spacing
async function customSpacing() {
  const doc = await Document.load('input.docx');

  const result = doc.ensureBlankLinesAfter1x1Tables({
    spacingAfter: 240,  // 12pt spacing instead of default 6pt
    markAsPreserved: true,
  });

  console.log(`Added/marked ${result.blankLinesAdded + result.existingLinesMarked} blank lines with 12pt spacing`);
  await doc.save('output.docx');
  doc.dispose();
}

// Run main example
main().catch(console.error);
