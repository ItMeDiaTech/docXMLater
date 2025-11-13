/**
 * Example: Multi-Level Lists
 *
 * Demonstrates complex multi-level lists with mixed numbering formats,
 * showing how to create professional documents with hierarchical information.
 */

import { Document } from '../../src';

async function demonstrateMultiLevelLists() {
  console.log('Creating document with multi-level lists...\n');

  const doc = Document.create({
    properties: {
      title: 'Multi-Level List Example',
      creator: 'DocXML',
    },
  });

  // Add title
  doc.createParagraph('Multi-Level Lists Example').setStyle('Title');
  doc.createParagraph('Advanced hierarchical list structures').setStyle('Subtitle');
  doc.createParagraph();

  // Introduction
  doc.createParagraph('Introduction').setStyle('Heading1');
  doc
    .createParagraph(
      'Multi-level lists allow you to create complex hierarchical structures ' +
        'with different numbering formats at each level. This is essential for ' +
        'formal documents, technical specifications, and structured content.'
    )
    .setStyle('Normal');
  doc.createParagraph();

  // Pre-configured multi-level list
  doc.createParagraph('Using createMultiLevelList()').setStyle('Heading1');
  doc
    .createParagraph(
      'The createMultiLevelList() method creates a list with automatic format progression:'
    )
    .setStyle('Normal');
  doc.createParagraph();

  const multiListId = doc.createMultiLevelList();

  doc.createParagraph('Level 0: Decimal (1, 2, 3...)').setNumbering(multiListId, 0);
  doc.createParagraph('Level 1: Lower letters (a, b, c...)').setNumbering(multiListId, 1);
  doc.createParagraph('Level 2: Lower roman (i, ii, iii...)').setNumbering(multiListId, 2);
  doc.createParagraph('Level 3: Decimal (1, 2, 3...)').setNumbering(multiListId, 3);
  doc.createParagraph('Another level 2 item').setNumbering(multiListId, 2);
  doc.createParagraph('Another level 1 item').setNumbering(multiListId, 1);
  doc.createParagraph('Level 0 continues').setNumbering(multiListId, 0);

  doc.createParagraph();

  // Corporate structure example
  doc.createParagraph('Corporate Structure Example').setStyle('Heading1');
  doc
    .createParagraph(
      'Here is how you might document a corporate organizational structure:'
    )
    .setStyle('Normal');
  doc.createParagraph();

  const corpListId = doc.createNumberedList(4, ['decimal', 'lowerLetter', 'lowerRoman', 'decimal']);

  doc.createParagraph('Executive Leadership').setNumbering(corpListId, 0);
  doc.createParagraph('Chief Executive Officer (CEO)').setNumbering(corpListId, 1);
  doc.createParagraph('Strategic planning').setNumbering(corpListId, 2);
  doc.createParagraph('Board relations').setNumbering(corpListId, 2);
  doc.createParagraph('Chief Technology Officer (CTO)').setNumbering(corpListId, 1);
  doc.createParagraph('Technology strategy').setNumbering(corpListId, 2);
  doc.createParagraph('Innovation initiatives').setNumbering(corpListId, 2);
  doc.createParagraph('Chief Financial Officer (CFO)').setNumbering(corpListId, 1);

  doc.createParagraph('Product Division').setNumbering(corpListId, 0);
  doc.createParagraph('Product Management').setNumbering(corpListId, 1);
  doc.createParagraph('Product roadmap').setNumbering(corpListId, 2);
  doc.createParagraph('User research').setNumbering(corpListId, 2);
  doc.createParagraph('Engineering').setNumbering(corpListId, 1);
  doc.createParagraph('Frontend team').setNumbering(corpListId, 2);
  doc.createParagraph('Backend team').setNumbering(corpListId, 2);
  doc.createParagraph('Infrastructure team').setNumbering(corpListId, 2);

  doc.createParagraph('Operations').setNumbering(corpListId, 0);
  doc.createParagraph('Human Resources').setNumbering(corpListId, 1);
  doc.createParagraph('Finance').setNumbering(corpListId, 1);
  doc.createParagraph('Facilities').setNumbering(corpListId, 1);

  doc.createParagraph();

  // Technical specification example
  doc.createParagraph('Technical Specification').setStyle('Heading1');
  doc
    .createParagraph(
      'Multi-level lists are ideal for technical specifications and requirements documents:'
    )
    .setStyle('Normal');
  doc.createParagraph();

  const specListId = doc.createNumberedList(4, ['decimal', 'decimal', 'lowerLetter', 'lowerRoman']);

  doc.createParagraph('System Requirements').setNumbering(specListId, 0);
  doc.createParagraph('Hardware Requirements').setNumbering(specListId, 1);
  doc.createParagraph('Processor: Intel Core i5 or equivalent').setNumbering(specListId, 2);
  doc.createParagraph('Memory: 8GB RAM minimum').setNumbering(specListId, 2);
  doc.createParagraph('Storage: 256GB SSD').setNumbering(specListId, 2);
  doc.createParagraph('Software Requirements').setNumbering(specListId, 1);
  doc.createParagraph('Operating System: Windows 10 or later').setNumbering(specListId, 2);
  doc.createParagraph('Runtime: Node.js 18+').setNumbering(specListId, 2);
  doc.createParagraph('Dependencies').setNumbering(specListId, 3);
  doc.createParagraph('Package A version 2.0+').setNumbering(specListId, 4);
  doc.createParagraph('Package B version 1.5+').setNumbering(specListId, 4);

  doc.createParagraph('Functional Requirements').setNumbering(specListId, 0);
  doc.createParagraph('User Authentication').setNumbering(specListId, 1);
  doc.createParagraph('Login functionality').setNumbering(specListId, 2);
  doc.createParagraph('Username/password').setNumbering(specListId, 3);
  doc.createParagraph('OAuth providers').setNumbering(specListId, 3);
  doc.createParagraph('Password reset').setNumbering(specListId, 2);
  doc.createParagraph('Two-factor authentication').setNumbering(specListId, 2);
  doc.createParagraph('Data Management').setNumbering(specListId, 1);
  doc.createParagraph('CRUD operations').setNumbering(specListId, 2);
  doc.createParagraph('Data validation').setNumbering(specListId, 2);
  doc.createParagraph('Backup and restore').setNumbering(specListId, 2);

  doc.createParagraph('Non-Functional Requirements').setNumbering(specListId, 0);
  doc.createParagraph('Performance').setNumbering(specListId, 1);
  doc.createParagraph('Page load time < 2 seconds').setNumbering(specListId, 2);
  doc.createParagraph('API response time < 500ms').setNumbering(specListId, 2);
  doc.createParagraph('Security').setNumbering(specListId, 1);
  doc.createParagraph('Data encryption at rest').setNumbering(specListId, 2);
  doc.createParagraph('HTTPS for all connections').setNumbering(specListId, 2);
  doc.createParagraph('Scalability').setNumbering(specListId, 1);

  doc.createParagraph();

  // Mixed bullets and numbers
  doc.createParagraph('Mixing Bullets and Numbers').setStyle('Heading1');
  doc
    .createParagraph(
      'You can create different types of lists in the same document:'
    )
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('Features (numbered):').setStyle('Heading2');
  const featureListId = doc.createNumberedList();
  doc.createParagraph('User authentication').setNumbering(featureListId, 0);
  doc.createParagraph('Data management').setNumbering(featureListId, 0);
  doc.createParagraph('Reporting').setNumbering(featureListId, 0);

  doc.createParagraph();
  doc.createParagraph('Technologies (bulleted):').setStyle('Heading2');
  const techListId = doc.createBulletList();
  doc.createParagraph('TypeScript').setNumbering(techListId, 0);
  doc.createParagraph('Node.js').setNumbering(techListId, 0);
  doc.createParagraph('React').setNumbering(techListId, 0);

  doc.createParagraph();
  doc.createParagraph('Installation Steps (numbered):').setStyle('Heading2');
  const installListId = doc.createNumberedList();
  doc.createParagraph('Clone the repository').setNumbering(installListId, 0);
  doc.createParagraph('Run npm install').setNumbering(installListId, 0);
  doc.createParagraph('Configure environment variables').setNumbering(installListId, 0);
  doc.createParagraph('Start the development server').setNumbering(installListId, 0);

  doc.createParagraph();

  // Summary
  doc.createParagraph('Summary').setStyle('Heading1');
  doc
    .createParagraph(
      'Multi-level lists are powerful tools for organizing complex information. ' +
        'DocXML makes it easy to create these structures with methods like ' +
        'createNumberedList(), createBulletList(), and createMultiLevelList(). ' +
        'Each paragraph can be assigned to any level (0-8) using setNumbering(numId, level).'
    )
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('Key Points:').setStyle('Heading2');
  const summaryListId = doc.createBulletList();
  doc.createParagraph('Lists support up to 9 levels (0-8)').setNumbering(summaryListId, 0);
  doc.createParagraph('Each level can have different numbering formats').setNumbering(summaryListId, 0);
  doc.createParagraph('Multiple lists can exist in the same document').setNumbering(summaryListId, 0);
  doc.createParagraph('Lists maintain proper numbering automatically').setNumbering(summaryListId, 0);

  // Save
  const filename = 'multi-level-list.docx';
  await doc.save(filename);
  console.log(`✓ Created ${filename}`);
  console.log('  Open in Microsoft Word to see complex multi-level lists!');
  console.log('\nLists created:');
  console.log('  • Pre-configured multi-level list');
  console.log('  • Corporate structure (4 levels)');
  console.log('  • Technical specification (5 levels)');
  console.log('  • Mixed bullet and numbered lists');
}

// Run the example
demonstrateMultiLevelLists().catch(console.error);
