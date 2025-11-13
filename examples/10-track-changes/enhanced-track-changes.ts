/**
 * Enhanced Track Changes Examples
 *
 * Demonstrates new tracked changes features:
 * - Settings.xml integration (w:trackRevisions flag)
 * - Range markers for move operations
 * - RSID tracking
 * - Document protection
 * - Revision view settings
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
 * Example 1: Enable track changes with settings
 */
async function example1_TrackChangesSettings() {
  console.log('Example 1: Track changes with settings.xml integration...');

  const doc = Document.create({
    properties: {
      title: 'Track Changes Settings Example',
      creator: 'DocXML Examples',
    },
  });

  // Enable track changes with custom settings
  doc.enableTrackChanges({
    trackFormatting: true,
    showInsertionsAndDeletions: true,
    showFormatting: true,
    showInkAnnotations: true,
  });

  doc.createParagraph('Track Changes with Settings')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'This document has track changes enabled in settings.xml. ' +
    'When you open it in Microsoft Word, the "Track Changes" button will be on, ' +
    'and all changes will be automatically tracked.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Add some tracked changes
  const para1 = doc.createParagraph('This is the original text. ');
  doc.trackInsertion(para1, 'Alice', 'This was added. ', new Date());
  doc.trackDeletion(para1, 'Bob', 'This was removed. ', new Date());
  para1.addText('Final text.')
    .setSpaceAfter(240);

  // Show settings info
  doc.createParagraph('Track Changes Settings:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  const settings = doc.getRevisionViewSettings();
  doc.createParagraph()
    .addText(`Enabled: ${doc.isTrackChangesEnabled()}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`Track Formatting: ${doc.isTrackFormattingEnabled()}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Show Ins/Del: ${settings.showInsertionsAndDeletions}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Show Formatting: ${settings.showFormatting}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'enhanced-example1-settings.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log('  Note: Track changes is enabled in settings.xml!');
}

/**
 * Example 2: Move operations with range markers
 */
async function example2_MoveWithRangeMarkers() {
  console.log('Example 2: Move operations with range markers...');

  const doc = Document.create({
    properties: {
      title: 'Move with Range Markers Example',
      creator: 'DocXML Examples',
    },
  });

  doc.enableTrackChanges();

  doc.createParagraph('Move Operations with Range Markers')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'Range markers properly mark the boundaries of moved content. ' +
    'This enables Word to correctly display and manage multi-paragraph moves.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Create source paragraph
  doc.createParagraph('Source Location:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  const sourcePara = doc.createParagraph();
  sourcePara.addText('The quick brown fox ');

  // Track a move operation
  const move = doc.trackMove('Alice', new Run('jumped over the lazy dog'), new Date());

  // Add range markers and revision to source
  sourcePara.addRangeMarker(move.moveFromRangeStart);
  sourcePara.addRevision(move.moveFrom);
  sourcePara.addRangeMarker(move.moveFromRangeEnd);

  sourcePara.addText(' and landed safely.')
    .setSpaceAfter(240);

  // Create destination paragraph
  doc.createParagraph('Destination Location:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  const destPara = doc.createParagraph();
  destPara.addText('The phrase "');

  // Add range markers and revision to destination
  destPara.addRangeMarker(move.moveToRangeStart);
  destPara.addRevision(move.moveTo);
  destPara.addRangeMarker(move.moveToRangeEnd);

  destPara.addText('" was moved here.')
    .setSpaceAfter(480);

  // Show move details
  doc.createParagraph('Move Details:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  doc.createParagraph()
    .addText(`Move ID: ${move.moveId}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`Range Markers: 4 (2 start, 2 end)`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Revisions: 2 (moveFrom, moveTo)`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'enhanced-example2-move-ranges.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log('  Note: Move operation uses proper range markers!');
}

/**
 * Example 3: RSID tracking
 */
async function example3_RsidTracking() {
  console.log('Example 3: RSID (Revision Save ID) tracking...');

  const doc = Document.create({
    properties: {
      title: 'RSID Tracking Example',
      creator: 'DocXML Examples',
    },
  });

  doc.enableTrackChanges();

  // Set up RSIDs
  doc.setRsidRoot('00A12B3C'); // First editing session
  doc.addRsid('00D45E6F');     // Second editing session
  doc.addRsid('00789ABC');     // Third editing session

  // Generate RSID for current session
  const currentRsid = doc.generateRsid();

  doc.createParagraph('RSID Tracking')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'RSIDs (Revision Save IDs) track editing sessions in the document. ' +
    'They help identify which changes were made in the same editing session, ' +
    'useful for document comparison and merge operations.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Show RSID information
  doc.createParagraph('Document RSIDs:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  doc.createParagraph()
    .addText(`RSID Root: ${doc.getRsidRoot()}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(60);

  const rsids = doc.getRsids();
  doc.createParagraph()
    .addText(`Total RSIDs: ${rsids.length}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  rsids.forEach((rsid, index) => {
    doc.createParagraph()
      .addText(`  ${index + 1}. ${rsid}`, { font: 'Courier New', size: 10 })
      .setSpaceBefore(20);
  });

  doc.createParagraph()
    .addText(`Current Session: ${currentRsid}`, { font: 'Courier New', size: 10, bold: true })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'enhanced-example3-rsid-tracking.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log(`  Note: RSIDs stored in settings.xml (${rsids.length} total)!`);
}

/**
 * Example 4: Document protection
 */
async function example4_DocumentProtection() {
  console.log('Example 4: Document protection with tracked changes...');

  const doc = Document.create({
    properties: {
      title: 'Protected Document Example',
      creator: 'DocXML Examples',
    },
  });

  doc.enableTrackChanges({
    trackFormatting: true,
    showInsertionsAndDeletions: true,
  });

  // Protect the document - force track changes
  doc.protectDocument({
    edit: 'trackedChanges',
    enforcement: true,
    password: 'secret123',
    cryptSpinCount: 100000,
  });

  doc.createParagraph('Protected Document with Track Changes')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'This document is protected with a password. All edits must be tracked changes. ' +
    'Users cannot disable track changes without the password (secret123).'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Add content
  const para1 = doc.createParagraph('Original content in protected document. ');
  doc.trackInsertion(para1, 'System', 'Added text (tracked). ', new Date());
  para1.addText('More content.')
    .setSpaceAfter(240);

  // Show protection info
  doc.createParagraph('Protection Details:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  const protection = doc.getProtection();
  doc.createParagraph()
    .addText(`Protected: ${doc.isProtected()}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`Edit Mode: ${protection?.edit}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Enforcement: ${protection?.enforcement}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Algorithm: PBKDF2 + SHA-512`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Spin Count: ${protection?.cryptSpinCount?.toLocaleString()}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'enhanced-example4-protected.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log('  Note: Document is password-protected (password: secret123)!');
}

/**
 * Example 5: Paragraph mark deletion
 */
async function example5_ParagraphMarkDeletion() {
  console.log('Example 5: Paragraph mark deletion tracking...');

  const doc = Document.create({
    properties: {
      title: 'Paragraph Mark Deletion Example',
      creator: 'DocXML Examples',
    },
  });

  doc.enableTrackChanges();

  doc.createParagraph('Paragraph Mark Deletion Tracking')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'When you delete a paragraph mark (¶) to join two paragraphs, ' +
    'Word tracks this as a paragraph mark deletion. ' +
    'The deletion appears in the paragraph properties (w:pPr/w:rPr/w:del).'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Example 1: Simple paragraph mark deletion
  doc.createParagraph('Simple Deletion:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  const para1 = doc.createParagraph('First paragraph');
  doc.trackParagraphMarkDeletion(para1, 'Alice', new Date());
  para1.addText(' (paragraph mark deleted - joined with next)');
  para1.setSpaceAfter(240);

  doc.createParagraph('This paragraph was originally separate but was joined.')
    .setSpaceAfter(240);

  // Example 2: Multiple deletions
  doc.createParagraph('Multiple Deletions:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  const para2 = doc.createParagraph('Paragraph A');
  doc.trackParagraphMarkDeletion(para2, 'Bob', new Date());

  const para3 = doc.createParagraph('Paragraph B');
  doc.trackParagraphMarkDeletion(para3, 'Carol', new Date());

  doc.createParagraph('Paragraph C (final)')
    .setSpaceAfter(240);

  doc.createParagraph(
    'Note: In Microsoft Word, open this document and show Track Changes. ' +
    'The ¶ (paragraph mark) symbols will appear as deleted, ' +
    'indicating where paragraphs were joined together.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Show deletion info
  doc.createParagraph('Deletion Details:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  doc.createParagraph()
    .addText(`Paragraph 1 mark deleted: ${para1.isParagraphMarkDeleted()}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`Paragraph 2 mark deleted: ${para2.isParagraphMarkDeleted()}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Paragraph 3 mark deleted: ${para3.isParagraphMarkDeleted()}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'enhanced-example5-paragraph-marks.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log('  Note: Paragraph marks tracked as deletions!');
}

/**
 * Example 6: Complete workflow
 */
async function example6_CompleteWorkflow() {
  console.log('Example 6: Complete track changes workflow...');

  const doc = Document.create({
    properties: {
      title: 'Complete Track Changes Workflow',
      creator: 'DocXML Examples',
    },
  });

  // Enable all features
  doc.enableTrackChanges({
    trackFormatting: true,
    showInsertionsAndDeletions: true,
    showFormatting: true,
  });

  doc.setRsidRoot('00ABC123');
  const sessionRsid = doc.generateRsid();

  doc.createParagraph('Complete Track Changes Workflow')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  // Section 1: Basic changes
  doc.createParagraph('1. Basic Text Changes')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const basic = doc.createParagraph('Original text ');
  doc.trackInsertion(basic, 'Alice', 'inserted ', new Date());
  basic.addText('more text ');
  doc.trackDeletion(basic, 'Bob', 'deleted ', new Date());
  basic.addText('final text.')
    .setSpaceAfter(240);

  // Section 2: Move operation
  doc.createParagraph('2. Move Operation')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const moveSrc = doc.createParagraph('Source: ');
  const move = doc.trackMove('Carol', new Run('moved content'), new Date());
  moveSrc.addRangeMarker(move.moveFromRangeStart);
  moveSrc.addRevision(move.moveFrom);
  moveSrc.addRangeMarker(move.moveFromRangeEnd);
  moveSrc.setSpaceAfter(120);

  const moveDest = doc.createParagraph('Destination: ');
  moveDest.addRangeMarker(move.moveToRangeStart);
  moveDest.addRevision(move.moveTo);
  moveDest.addRangeMarker(move.moveToRangeEnd);
  moveDest.setSpaceAfter(240);

  // Section 3: Formatting changes
  doc.createParagraph('3. Formatting Changes')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const fmt = doc.createParagraph('Text with ');
  const boldRun = new Run('bold formatting', { bold: true });
  const fmtChange = doc.createRunPropertiesChange('Dave', boldRun, { bold: false });
  fmt.addRevision(fmtChange);
  fmt.addText(' applied.')
    .setSpaceAfter(480);

  // Summary
  doc.createParagraph('Summary')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const stats = doc.getRevisionStats();
  doc.createParagraph()
    .addText(`Total Revisions: ${stats.total}`, { font: 'Courier New' })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`RSIDs: ${doc.getRsids().length}`, { font: 'Courier New' })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Current Session: ${sessionRsid}`, { font: 'Courier New' })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'enhanced-example5-complete.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log('  Note: Complete workflow with all features enabled!');
}

/**
 * Main function to run all examples
 */
async function main() {
  console.log('Running Enhanced Track Changes Examples...\n');

  try {
    await example1_TrackChangesSettings();
    await example2_MoveWithRangeMarkers();
    await example3_RsidTracking();
    await example4_DocumentProtection();
    await example5_ParagraphMarkDeletion();
    await example6_CompleteWorkflow();

    console.log('\n✓ All examples completed successfully!');
    console.log(`\nOutput files saved to: ${outputDir}`);
    console.log('\nNew Features Demonstrated:');
    console.log('  ✓ Track changes enabled in settings.xml');
    console.log('  ✓ Range markers for move operations');
    console.log('  ✓ RSID (Revision Save ID) tracking');
    console.log('  ✓ Document protection with passwords');
    console.log('  ✓ Paragraph mark deletion tracking');
    console.log('  ✓ Revision view settings');
    console.log('\nOpen documents in Microsoft Word to see track changes working!');
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
  example1_TrackChangesSettings,
  example2_MoveWithRangeMarkers,
  example3_RsidTracking,
  example4_DocumentProtection,
  example5_ParagraphMarkDeletion,
  example6_CompleteWorkflow,
};
