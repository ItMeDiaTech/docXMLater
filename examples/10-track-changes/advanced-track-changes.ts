/**
 * Advanced Track Changes Examples
 *
 * Demonstrates all types of tracked changes supported by Microsoft Word:
 * - Content changes (insert, delete)
 * - Property changes (formatting changes)
 * - Move operations (cut and paste)
 * - Table cell operations
 * - Numbering changes
 */

import { Document, Run } from '../../src';
import * as path from 'path';
import * as fs from 'fs';

// Create output directory if it doesn't exist
const outputDir = __dirname;
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

/**
 * Example 1: Property Changes - Track formatting changes
 */
async function example1_PropertyChanges() {
  console.log('Example 1: Property changes (formatting revisions)...');

  const doc = Document.create({
    properties: {
      title: 'Property Changes Example',
      creator: 'DocXML Examples',
    },
  });

  doc.createParagraph('Track Changes: Property Modifications')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'This document demonstrates tracking of formatting changes. When you open this in Microsoft Word, ' +
    'you will see tracked changes for text formatting modifications.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Run Properties Change: Bold formatting added
  doc.createParagraph('Run Property Changes')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const para1 = doc.createParagraph();
  para1.addText('This text was originally plain, but ');

  // Create a run with new bold formatting and track the change
  const boldRun = new Run('this part became bold', { bold: true });
  const boldChange = doc.createRunPropertiesChange(
    'Alice',
    boldRun,
    { bold: false }, // Previous state: not bold
    new Date('2025-10-16T10:00:00Z')
  );
  para1.addRevision(boldChange);

  para1.addText(', and ');

  // Create a run with italic and color, tracking the change
  const styledRun = new Run('this part got italic and red color', { italic: true, color: 'FF0000' });
  const styleChange = doc.createRunPropertiesChange(
    'Bob',
    styledRun,
    { italic: false, color: '000000' }, // Previous state
    new Date('2025-10-16T10:30:00Z')
  );
  para1.addRevision(styleChange);

  para1.addText('.')
    .setSpaceAfter(240);

  // Paragraph Properties Change: Alignment changed
  doc.createParagraph('Paragraph Property Changes')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const para2 = doc.createParagraph();
  para2.addText('This paragraph had its alignment changed from left to center.');

  // Track the paragraph property change
  const alignmentChange = doc.createParagraphPropertiesChange(
    'Carol',
    new Run(''),
    { alignment: 'left' }, // Previous alignment
    new Date('2025-10-16T11:00:00Z')
  );
  para2.addRevision(alignmentChange);
  para2.setAlignment('center')
    .setSpaceAfter(480);

  // Statistics
  const stats = doc.getRevisionStats();
  doc.createParagraph('Statistics:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  doc.createParagraph()
    .addText(`Total revisions: ${stats.total}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`Property changes: ${stats.propertyChanges}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'example1-property-changes.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 2: Move Operations - Track cut and paste
 */
async function example2_MoveOperations() {
  console.log('Example 2: Move operations (cut and paste tracking)...');

  const doc = Document.create({
    properties: {
      title: 'Move Operations Example',
      creator: 'DocXML Examples',
    },
  });

  doc.createParagraph('Track Changes: Move Operations')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'Move operations track when content is cut from one location and pasted to another. ' +
    'Microsoft Word will show the source and destination of moved content.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Original paragraph before move
  doc.createParagraph('Original Content:')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const originalPara = doc.createParagraph();
  originalPara.addText('The quick brown fox ')
    .setAlignment('justify')
    .setSpaceAfter(240);

  // After move: Show moveFrom and moveTo
  doc.createParagraph('After Move:')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const afterMove1 = doc.createParagraph();
  afterMove1.addText('The quick brown fox ')
    .setAlignment('justify');

  // Track the move operation
  const movedText = new Run('jumped over the lazy dog');
  const moveOp = doc.trackMove('Alice', movedText, new Date('2025-10-16T10:00:00Z'));

  // Add moveFrom (original location)
  afterMove1.addRevision(moveOp.moveFrom);
  afterMove1.addText(' and landed safely.')
    .setSpaceAfter(240);

  // Add moveTo (new location)
  const afterMove2 = doc.createParagraph();
  afterMove2.addText('The phrase "');
  afterMove2.addRevision(moveOp.moveTo);
  afterMove2.addText('" was moved here from the previous paragraph.')
    .setAlignment('justify')
    .setSpaceAfter(480);

  // Multiple moves
  doc.createParagraph('Multiple Moves:')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const move1 = doc.trackMove('Bob', new Run('first phrase'), new Date('2025-10-16T11:00:00Z'));
  const move2 = doc.trackMove('Bob', new Run('second phrase'), new Date('2025-10-16T11:30:00Z'));

  const movePara1 = doc.createParagraph();
  movePara1.addText('Original: ');
  movePara1.addRevision(move1.moveFrom);
  movePara1.addText(' and ');
  movePara1.addRevision(move2.moveFrom);
  movePara1.addText('.')
    .setSpaceAfter(120);

  const movePara2 = doc.createParagraph();
  movePara2.addText('After rearranging: ');
  movePara2.addRevision(move2.moveTo);
  movePara2.addText(' and ');
  movePara2.addRevision(move1.moveTo);
  movePara2.addText('.')
    .setSpaceAfter(480);

  // Statistics
  const stats = doc.getRevisionStats();
  doc.createParagraph('Statistics:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  doc.createParagraph()
    .addText(`Total revisions: ${stats.total}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`Move operations: ${stats.moves}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'example2-move-operations.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 3: Table Cell Operations
 */
async function example3_TableCellOperations() {
  console.log('Example 3: Table cell operations...');

  const doc = Document.create({
    properties: {
      title: 'Table Cell Operations Example',
      creator: 'DocXML Examples',
    },
  });

  doc.createParagraph('Track Changes: Table Operations')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'This demonstrates tracking of table cell insertions, deletions, and merges.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Table Cell Insert
  doc.createParagraph('Cell Insertion:')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const insertPara = doc.createParagraph();
  insertPara.addText('A new cell was inserted: ');
  const cellInsert = doc.createTableCellInsert(
    'Alice',
    new Run('New Cell Content'),
    new Date('2025-10-16T10:00:00Z')
  );
  insertPara.addRevision(cellInsert);
  insertPara.setSpaceAfter(240);

  // Table Cell Delete
  doc.createParagraph('Cell Deletion:')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const deletePara = doc.createParagraph();
  deletePara.addText('A cell was deleted: ');
  const cellDelete = doc.createTableCellDelete(
    'Bob',
    new Run('Deleted Cell Content'),
    new Date('2025-10-16T11:00:00Z')
  );
  deletePara.addRevision(cellDelete);
  deletePara.setSpaceAfter(240);

  // Table Cell Merge
  doc.createParagraph('Cell Merge:')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const mergePara = doc.createParagraph();
  mergePara.addText('Two cells were merged: ');
  const cellMerge = doc.createTableCellMerge(
    'Carol',
    new Run('Merged Cell Content'),
    new Date('2025-10-16T12:00:00Z')
  );
  mergePara.addRevision(cellMerge);
  mergePara.setSpaceAfter(480);

  // Statistics
  const stats = doc.getRevisionStats();
  doc.createParagraph('Statistics:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  doc.createParagraph()
    .addText(`Total revisions: ${stats.total}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`Table cell changes: ${stats.tableCellChanges}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'example3-table-cell-operations.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 4: Numbering Changes
 */
async function example4_NumberingChanges() {
  console.log('Example 4: Numbering changes...');

  const doc = Document.create({
    properties: {
      title: 'Numbering Changes Example',
      creator: 'DocXML Examples',
    },
  });

  doc.createParagraph('Track Changes: Numbering Modifications')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'This demonstrates tracking of list numbering format changes.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Numbering Format Change:')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const numberingPara = doc.createParagraph();
  numberingPara.addText('This list item had its numbering format changed: ');

  // Track numbering change
  const numberingChange = doc.createNumberingChange(
    'Alice',
    new Run('Important item'),
    { numId: 1, ilvl: 0, format: 'decimal' }, // Previous numbering
    new Date('2025-10-16T10:00:00Z')
  );
  numberingPara.addRevision(numberingChange);
  numberingPara.setSpaceAfter(480);

  // Statistics
  const stats = doc.getRevisionStats();
  doc.createParagraph('Statistics:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  doc.createParagraph()
    .addText(`Total revisions: ${stats.total}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`Property changes: ${stats.propertyChanges}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'example4-numbering-changes.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 5: Comprehensive - All revision types in one document
 */
async function example5_AllRevisionTypes() {
  console.log('Example 5: All revision types in one document...');

  const doc = Document.create({
    properties: {
      title: 'All Revision Types',
      creator: 'DocXML Examples',
    },
  });

  doc.createParagraph('Complete Track Changes Demo')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'This document contains examples of all supported tracked change types in Microsoft Word.'
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  // 1. Content Changes
  doc.createParagraph('1. Content Changes')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const contentPara = doc.createParagraph();
  contentPara.addText('Original text ');
  doc.trackInsertion(contentPara, 'Alice', 'inserted text ', new Date('2025-10-16T09:00:00Z'));
  contentPara.addText('more text ');
  doc.trackDeletion(contentPara, 'Bob', 'deleted text ', new Date('2025-10-16T09:30:00Z'));
  contentPara.addText('final text.')
    .setSpaceAfter(240);

  // 2. Property Changes
  doc.createParagraph('2. Formatting Changes')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const propPara = doc.createParagraph();
  propPara.addText('Text with ');
  const boldRun = new Run('bold formatting', { bold: true });
  const boldChange = doc.createRunPropertiesChange('Carol', boldRun, { bold: false });
  propPara.addRevision(boldChange);
  propPara.addText(' tracked.')
    .setSpaceAfter(240);

  // 3. Move Operations
  doc.createParagraph('3. Move Operations')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const movePara1 = doc.createParagraph();
  movePara1.addText('Source: ');
  const moveOp = doc.trackMove('Dave', new Run('moved content'), new Date('2025-10-16T10:00:00Z'));
  movePara1.addRevision(moveOp.moveFrom);
  movePara1.setSpaceAfter(120);

  const movePara2 = doc.createParagraph();
  movePara2.addText('Destination: ');
  movePara2.addRevision(moveOp.moveTo);
  movePara2.setSpaceAfter(240);

  // 4. Table Operations
  doc.createParagraph('4. Table Cell Changes')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const tablePara = doc.createParagraph();
  tablePara.addText('Cell inserted: ');
  const cellInsert = doc.createTableCellInsert('Eve', new Run('New cell'));
  tablePara.addRevision(cellInsert);
  tablePara.setSpaceAfter(240);

  // 5. Numbering Changes
  doc.createParagraph('5. Numbering Changes')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const numPara = doc.createParagraph();
  numPara.addText('Numbering format changed: ');
  const numChange = doc.createNumberingChange('Frank', new Run('List item'), { numId: 1 });
  numPara.addRevision(numChange);
  numPara.setSpaceAfter(480);

  // Complete Statistics
  const stats = doc.getRevisionStats();
  doc.createParagraph('Complete Statistics')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  doc.createParagraph()
    .addText(`Total revisions: ${stats.total}`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`Insertions: ${stats.insertions}`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Deletions: ${stats.deletions}`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Property changes: ${stats.propertyChanges}`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Move operations: ${stats.moves}`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Table cell changes: ${stats.tableCellChanges}`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Authors: ${stats.authors.join(', ')}`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'example5-all-revision-types.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Main function to run all examples
 */
async function main() {
  console.log('Running Advanced Track Changes Examples...\n');

  try {
    await example1_PropertyChanges();
    await example2_MoveOperations();
    await example3_TableCellOperations();
    await example4_NumberingChanges();
    await example5_AllRevisionTypes();

    console.log('\n✓ All examples completed successfully!');
    console.log(`\nOutput files saved to: ${outputDir}`);
    console.log('\nOpen these documents in Microsoft Word to see:');
    console.log('   - Content changes (insertions and deletions)');
    console.log('   - Formatting changes (property modifications)');
    console.log('   - Move operations (cut and paste tracking)');
    console.log('   - Table cell operations (insert, delete, merge)');
    console.log('   - Numbering changes (list format modifications)');
    console.log('\nAll changes are tracked with author names and timestamps!');
  } catch (error) {
    console.error('Error running examples:', error);
    process.exit(1);
  }
}

// Run examples if executed directly
if (require.main === module) {
  main();
}

export {
  example1_PropertyChanges,
  example2_MoveOperations,
  example3_TableCellOperations,
  example4_NumberingChanges,
  example5_AllRevisionTypes,
};
