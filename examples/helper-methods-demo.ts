/**
 * Demo script for new Document helper methods
 *
 * Demonstrates:
 * 1. getAllRuns() - Get all text runs from document
 * 2. removeFormattingFromAll() - Remove specific formatting from all runs
 * 3. updateAllHyperlinks() - Apply formatter function to all hyperlinks
 */

import { Document } from '../src/core/Document';
import { Hyperlink } from '../src/elements/Hyperlink';

async function demo() {
  console.log('=== Document Helper Methods Demo ===\n');

  // Create a test document with various formatting
  const doc = Document.create();

  // Add paragraphs with different formatting
  const para1 = doc.createParagraph();
  para1.addText('This is bold text', { bold: true });
  para1.addText(' and this is italic', { italic: true });
  para1.addText(' and this is highlighted', { highlight: 'yellow' });

  const para2 = doc.createParagraph();
  para2.addText('More text with ', { color: 'FF0000' });
  para2.addText('different colors', { color: '0000FF' });
  para2.addText(' and underline', { underline: true });

  // Add hyperlinks
  const para3 = doc.createParagraph('Visit our sites: ');
  para3.addHyperlink(Hyperlink.createExternal('https://example.com', 'Example'));
  para3.addText(' and ');
  para3.addHyperlink(Hyperlink.createExternal('https://internal.company.com', 'Internal Site'));

  // Add a table with formatted content
  const table = doc.createTable(2, 2);
  const cell = table.getRow(0)?.getCell(0);
  if (cell) {
    const cellPara = cell.getParagraphs()[0];
    if (cellPara) {
      cellPara.addText('Table text', { bold: true, italic: true });
    }
  }

  console.log('Document created with formatted content\n');

  // ========================================
  // DEMO 1: getAllRuns()
  // ========================================
  console.log('--- DEMO 1: getAllRuns() ---');
  const allRuns = doc.getAllRuns();
  console.log(`Total runs in document: ${allRuns.length}`);
  console.log('Runs:');
  allRuns.forEach((run, index) => {
    const text = run.getText();
    const formatting = run.getFormatting();
    const formatDetails: string[] = [];
    if (formatting.bold) formatDetails.push('bold');
    if (formatting.italic) formatDetails.push('italic');
    if (formatting.underline) formatDetails.push('underline');
    if (formatting.color) formatDetails.push(`color:${formatting.color}`);
    if (formatting.highlight) formatDetails.push(`highlight:${formatting.highlight}`);

    console.log(`  ${index + 1}. "${text}" [${formatDetails.join(', ') || 'no formatting'}]`);
  });
  console.log();

  // ========================================
  // DEMO 2: removeFormattingFromAll()
  // ========================================
  console.log('--- DEMO 2: removeFormattingFromAll() ---');

  // Count bold runs before removal
  let boldCount = allRuns.filter(run => run.getFormatting().bold).length;
  console.log(`Runs with bold BEFORE removal: ${boldCount}`);

  // Remove all bold formatting
  const removedCount = doc.removeFormattingFromAll('bold');
  console.log(`Removed bold from ${removedCount} runs`);

  // Count bold runs after removal
  const allRunsAfter = doc.getAllRuns();
  boldCount = allRunsAfter.filter(run => run.getFormatting().bold).length;
  console.log(`Runs with bold AFTER removal: ${boldCount}`);
  console.log();

  // Remove highlight from all runs
  const highlightRemoved = doc.removeFormattingFromAll('highlight');
  console.log(`Removed highlight from ${highlightRemoved} runs\n`);

  // ========================================
  // DEMO 3: updateAllHyperlinks()
  // ========================================
  console.log('--- DEMO 3: updateAllHyperlinks() ---');

  const hyperlinks = doc.getHyperlinks();
  console.log(`Hyperlinks in document: ${hyperlinks.length}`);

  console.log('BEFORE formatting:');
  hyperlinks.forEach(({ hyperlink }, index) => {
    const fmt = hyperlink.getFormatting();
    console.log(`  ${index + 1}. "${hyperlink.getText()}" - ${hyperlink.getUrl()}`);
    console.log(`     Formatting: color=${fmt.color}, bold=${fmt.bold}, underline=${fmt.underline}`);
  });

  // Apply conditional formatting based on URL
  const updatedCount = doc.updateAllHyperlinks((link) => {
    const url = link.getUrl();
    if (url?.includes('internal')) {
      // Internal links: blue and bold
      link.setFormatting({ color: '0000FF', bold: true, underline: 'single' });
    } else {
      // External links: red and italic
      link.setFormatting({ color: 'FF0000', italic: true, underline: 'single' });
    }
  });

  console.log(`\nUpdated ${updatedCount} hyperlinks`);

  console.log('\nAFTER formatting:');
  const hyperlinksAfter = doc.getHyperlinks();
  hyperlinksAfter.forEach(({ hyperlink }, index) => {
    const fmt = hyperlink.getFormatting();
    console.log(`  ${index + 1}. "${hyperlink.getText()}" - ${hyperlink.getUrl()}`);
    console.log(`     Formatting: color=${fmt.color}, bold=${fmt.bold}, italic=${fmt.italic}`);
  });
  console.log();

  // ========================================
  // Save the document
  // ========================================
  await doc.save('examples/output/helper-methods-demo.docx');
  console.log('Document saved to: examples/output/helper-methods-demo.docx');
  console.log('\n=== Demo Complete ===');
}

// Run the demo
demo().catch(console.error);
