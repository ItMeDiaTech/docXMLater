/**
 * Track Changes Usage Examples
 *
 * Demonstrates how to use track changes (revision tracking) in documents.
 * Track changes shows insertions, deletions, and who made each change.
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
 * Example 1: Simple insertions and deletions
 */
async function example1_BasicTrackChanges() {
  console.log('Example 1: Basic track changes (insertions and deletions)...');

  const doc = Document.create({
    properties: {
      title: 'Basic Track Changes Example',
      creator: 'DocXML Examples',
    },
  });

  doc.createParagraph('Track Changes Demo')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'This document demonstrates track changes. Insertions appear with underlines ' +
    'and deletions appear with strikethroughs when you open it in Microsoft Word.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Create a paragraph with mixed content and tracked changes
  const para1 = doc.createParagraph();
  para1.addText('This is the original text. ');

  // Track an insertion by Alice
  const insertion1 = doc.createRevisionFromText(
    'insert',
    'Alice',
    'This text was added by Alice. ',
    new Date('2025-10-16T10:00:00Z')
  );
  para1.addRevision(insertion1);

  para1.addText('More original text. ');

  // Track a deletion by Bob
  const deletion1 = doc.createRevisionFromText(
    'delete',
    'Bob',
    'This text was deleted by Bob. ',
    new Date('2025-10-16T11:00:00Z')
  );
  para1.addRevision(deletion1);

  para1.addText('Final original text.')
    .setSpaceAfter(240);

  // Another paragraph with tracked changes
  const para2 = doc.createParagraph();
  para2.addText('Second paragraph with ');

  const insertion2 = doc.createRevisionFromText(
    'insert',
    'Alice',
    'inserted content',
    new Date('2025-10-16T12:00:00Z')
  );
  para2.addRevision(insertion2);

  para2.addText(' and ');

  const deletion2 = doc.createRevisionFromText(
    'delete',
    'Bob',
    'deleted content',
    new Date('2025-10-16T13:00:00Z')
  );
  para2.addRevision(deletion2);

  para2.addText(' in one line.')
    .setSpaceAfter(480);

  // Show revision statistics
  const stats = doc.getRevisionStats();
  doc.createParagraph('Revision Statistics:')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  doc.createParagraph()
    .addText(`Total revisions: ${stats.total}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`Insertions: ${stats.insertions}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Deletions: ${stats.deletions}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Authors: ${stats.authors.join(', ')}`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'example1-basic-track-changes.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log('  Note: Open in Word to see insertions (underlined) and deletions (strikethrough)!');
}

/**
 * Example 2: Collaborative editing with multiple authors
 */
async function example2_MultipleAuthors() {
  console.log('Example 2: Multiple authors making changes...');

  const doc = Document.create({
    properties: {
      title: 'Collaborative Editing Example',
      creator: 'DocXML Examples',
    },
  });

  doc.createParagraph('Collaborative Document Editing')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'This document shows how multiple people can make tracked changes. ' +
    'Each author\'s changes are tracked separately with their name and timestamp.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Original draft by Editor
  doc.createParagraph('Version 1: Original Draft')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  doc.createParagraph(
    'The project began in 2024 with a small team. The initial goals were modest. ' +
    'After six months, we had achieved significant progress.'
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  // Version 2: Alice's revisions
  doc.createParagraph('Version 2: Alice\'s Revisions (Morning)')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const alice1 = doc.createParagraph();
  alice1.addText('The project began in ');

  doc.trackDeletion(alice1, 'Alice', '2024', new Date('2025-10-16T09:00:00Z'));
  doc.trackInsertion(alice1, 'Alice', 'early 2025', new Date('2025-10-16T09:01:00Z'));

  alice1.addText(' with a small team. The initial goals were modest. ' +
    'After six months, we had achieved ');

  doc.trackDeletion(alice1, 'Alice', 'significant', new Date('2025-10-16T09:02:00Z'));
  doc.trackInsertion(alice1, 'Alice', 'remarkable', new Date('2025-10-16T09:03:00Z'));

  alice1.addText(' progress.')
    .setAlignment('justify')
    .setSpaceAfter(480);

  // Version 3: Bob's revisions
  doc.createParagraph('Version 3: Bob\'s Additional Revisions (Afternoon)')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const bob1 = doc.createParagraph();
  bob1.addText('The project began in ');

  doc.trackDeletion(bob1, 'Alice', '2024', new Date('2025-10-16T09:00:00Z'));
  doc.trackInsertion(bob1, 'Alice', 'early 2025', new Date('2025-10-16T09:01:00Z'));

  bob1.addText(' with a small ');

  doc.trackDeletion(bob1, 'Bob', 'team', new Date('2025-10-16T14:00:00Z'));
  doc.trackInsertion(bob1, 'Bob', 'but dedicated team of five people', new Date('2025-10-16T14:01:00Z'));

  bob1.addText('. The initial goals were modest. ' +
    'After six months, we had achieved ');

  doc.trackDeletion(bob1, 'Alice', 'significant', new Date('2025-10-16T09:02:00Z'));
  doc.trackInsertion(bob1, 'Alice', 'remarkable', new Date('2025-10-16T09:03:00Z'));

  bob1.addText(' progress.')
    .setAlignment('justify')
    .setSpaceAfter(480);

  // Version 4: Carol's revisions
  doc.createParagraph('Version 4: Carol\'s Final Revisions (Evening)')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const carol1 = doc.createParagraph();
  carol1.addText('The project began in ');

  doc.trackDeletion(carol1, 'Alice', '2024', new Date('2025-10-16T09:00:00Z'));
  doc.trackInsertion(carol1, 'Alice', 'early 2025', new Date('2025-10-16T09:01:00Z'));

  carol1.addText(' with a small ');

  doc.trackDeletion(carol1, 'Bob', 'team', new Date('2025-10-16T14:00:00Z'));
  doc.trackInsertion(carol1, 'Bob', 'but dedicated team of five people', new Date('2025-10-16T14:01:00Z'));

  carol1.addText('. ');

  doc.trackDeletion(carol1, 'Carol', 'The initial goals were modest. ', new Date('2025-10-16T18:00:00Z'));
  doc.trackInsertion(carol1, 'Carol', 'We set ambitious but achievable goals. ', new Date('2025-10-16T18:01:00Z'));

  carol1.addText('After six months, we had achieved ');

  doc.trackDeletion(carol1, 'Alice', 'significant', new Date('2025-10-16T09:02:00Z'));
  doc.trackInsertion(carol1, 'Alice', 'remarkable', new Date('2025-10-16T09:03:00Z'));

  carol1.addText(' progress.')
    .setAlignment('justify')
    .setSpaceAfter(480);

  // Summary
  const stats = doc.getRevisionStats();
  doc.createParagraph('Revision Summary')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  doc.createParagraph()
    .addText(`Total changes: ${stats.total}`, { font: 'Courier New' })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`Contributors: ${stats.authors.join(', ')}`, { font: 'Courier New' })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'example2-multiple-authors.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 3: Document review process with tracked changes
 */
async function example3_DocumentReview() {
  console.log('Example 3: Document review with track changes...');

  const doc = Document.create({
    properties: {
      title: 'Document Review Example',
      creator: 'DocXML Examples',
    },
  });

  doc.createParagraph('Document Review Process')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  // Section 1: Executive Summary
  doc.createParagraph('1. Executive Summary')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const summary = doc.createParagraph();
  summary.addText('This report presents our findings from the ');

  doc.trackInsertion(summary, 'Reviewer', 'comprehensive ', new Date('2025-10-15T10:00:00Z'));

  summary.addText('Q3 analysis. We ');

  doc.trackDeletion(summary, 'Reviewer', 'believe that ', new Date('2025-10-15T10:05:00Z'));
  doc.trackInsertion(summary, 'Reviewer', 'recommend ', new Date('2025-10-15T10:06:00Z'));

  summary.addText('the company should ');

  doc.trackDeletion(summary, 'Reviewer', 'consider ', new Date('2025-10-15T10:10:00Z'));

  summary.addText('pursue the new market opportunities ');

  doc.trackInsertion(summary, 'Reviewer', 'identified in this report', new Date('2025-10-15T10:15:00Z'));

  summary.addText('.')
    .setAlignment('justify')
    .setSpaceAfter(480);

  // Section 2: Methodology
  doc.createParagraph('2. Methodology')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const methodology = doc.createParagraph();
  methodology.addText('We ');

  doc.trackDeletion(methodology, 'Reviewer', 'used ', new Date('2025-10-15T11:00:00Z'));
  doc.trackInsertion(methodology, 'Reviewer', 'employed ', new Date('2025-10-15T11:01:00Z'));

  methodology.addText('a mixed-methods approach, combining ');

  doc.trackInsertion(methodology, 'Reviewer', 'both ', new Date('2025-10-15T11:05:00Z'));

  methodology.addText('quantitative ');

  doc.trackDeletion(methodology, 'Reviewer', 'data ', new Date('2025-10-15T11:08:00Z'));
  doc.trackInsertion(methodology, 'Reviewer', 'analysis ', new Date('2025-10-15T11:09:00Z'));

  methodology.addText('and qualitative research.');

  doc.trackDeletion(methodology, 'Reviewer', ' The data was collected over three months.', new Date('2025-10-15T11:12:00Z'));

  methodology
    .setAlignment('justify')
    .setSpaceAfter(480);

  // Section 3: Findings
  doc.createParagraph('3. Key Findings')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  const findings = doc.createParagraph();
  findings.addText('Our research ');

  doc.trackDeletion(findings, 'Reviewer', 'shows ', new Date('2025-10-15T12:00:00Z'));
  doc.trackInsertion(findings, 'Reviewer', 'demonstrates ', new Date('2025-10-15T12:01:00Z'));

  findings.addText('that customer satisfaction has ');

  doc.trackDeletion(findings, 'Reviewer', 'increased', new Date('2025-10-15T12:05:00Z'));
  doc.trackInsertion(findings, 'Reviewer', 'risen significantly', new Date('2025-10-15T12:06:00Z'));

  findings.addText(' over the ');

  doc.trackInsertion(findings, 'Reviewer', 'past ', new Date('2025-10-15T12:08:00Z'));

  findings.addText('quarter. ');

  doc.trackInsertion(findings, 'Reviewer', 'Specifically, our Net Promoter Score improved from 45 to 68. ', new Date('2025-10-15T12:10:00Z'));

  findings.addText('This trend ');

  doc.trackDeletion(findings, 'Reviewer', 'suggests ', new Date('2025-10-15T12:15:00Z'));
  doc.trackInsertion(findings, 'Reviewer', 'indicates ', new Date('2025-10-15T12:16:00Z'));

  findings.addText('strong market positioning.')
    .setAlignment('justify')
    .setSpaceAfter(480);

  // Show review statistics
  const stats = doc.getRevisionStats();
  doc.createParagraph('Review Statistics')
    .setStyle('Heading1')
    .setSpaceAfter(120);

  doc.createParagraph()
    .addText(`Total edits: ${stats.total}`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText(`Text added: ${stats.insertions} insertions`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Text removed: ${stats.deletions} deletions`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(40);

  doc.createParagraph()
    .addText(`Reviewer: ${stats.authors[0]}`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(40);

  // Save document
  const outputPath = path.join(outputDir, 'example3-document-review.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 4: Using Revision objects directly with formatted runs
 */
async function example4_FormattedRevisions() {
  console.log('Example 4: Formatted text in revisions...');

  const doc = Document.create({
    properties: {
      title: 'Formatted Revisions Example',
      creator: 'DocXML Examples',
    },
  });

  doc.createParagraph('Formatted Track Changes')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'This example shows how tracked changes can include formatted text ' +
    '(bold, italic, colors, etc.).'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Create paragraph with formatted tracked changes
  const para1 = doc.createParagraph();
  para1.addText('This paragraph contains ');

  // Create a formatted run for insertion
  const boldRun = new Run('bold inserted text', { bold: true });
  const insertionRev = doc.createInsertion('Author', boldRun, new Date('2025-10-16T10:00:00Z'));
  para1.addRevision(insertionRev);

  para1.addText(' and ');

  // Create a formatted run for deletion
  const italicRun = new Run('italic deleted text', { italic: true });
  const deletionRev = doc.createDeletion('Author', italicRun, new Date('2025-10-16T11:00:00Z'));
  para1.addRevision(deletionRev);

  para1.addText(' with formatting.')
    .setSpaceAfter(240);

  // Multiple formatted runs in one revision
  const para2 = doc.createParagraph();
  para2.addText('Complex revision with ');

  const runs = [
    new Run('bold', { bold: true }),
    new Run(', '),
    new Run('italic', { italic: true }),
    new Run(', and '),
    new Run('colored text', { color: 'FF0000' }),
  ];

  const complexInsertion = doc.createInsertion('Author', runs, new Date('2025-10-16T12:00:00Z'));
  para2.addRevision(complexInsertion);

  para2.addText(' in one tracked change.')
    .setSpaceAfter(240);

  // Save document
  const outputPath = path.join(outputDir, 'example4-formatted-revisions.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Main function to run all examples
 */
async function main() {
  console.log('Running Track Changes Examples...\n');

  try {
    await example1_BasicTrackChanges();
    await example2_MultipleAuthors();
    await example3_DocumentReview();
    await example4_FormattedRevisions();

    console.log('\n✓ All examples completed successfully!');
    console.log(`\nOutput files saved to: ${outputDir}`);
    console.log('\n📝 Important: Open the documents in Microsoft Word to see:');
    console.log('   - Insertions (usually shown with underlines)');
    console.log('   - Deletions (usually shown with strikethroughs)');
    console.log('   - Author names and timestamps for each change');
    console.log('\n   You can accept/reject changes using the Review tab in Word!');
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
  example1_BasicTrackChanges,
  example2_MultipleAuthors,
  example3_DocumentReview,
  example4_FormattedRevisions,
};
