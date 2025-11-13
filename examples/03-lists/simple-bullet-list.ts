/**
 * Example: Simple Bullet List
 *
 * Demonstrates how to create simple bullet lists with DocXML.
 */

import { Document } from '../../src';

async function demonstrateBulletList() {
  console.log('Creating document with bullet lists...\n');

  const doc = Document.create({
    properties: {
      title: 'Bullet List Example',
      creator: 'DocXML',
    },
  });

  // Add title
  doc.createParagraph('Simple Bullet List Example').setStyle('Title');
  doc.createParagraph();

  // Introduction
  doc.createParagraph('Introduction').setStyle('Heading1');
  doc
    .createParagraph(
      'This document demonstrates how to create bullet lists using DocXML.'
    )
    .setStyle('Normal');
  doc.createParagraph();

  // Create a bullet list
  doc.createParagraph('Shopping List').setStyle('Heading2');

  // Create bullet list
  const bulletListId = doc.createBulletList();

  // Add list items
  doc.createParagraph('Milk').setNumbering(bulletListId, 0);
  doc.createParagraph('Eggs').setNumbering(bulletListId, 0);
  doc.createParagraph('Bread').setNumbering(bulletListId, 0);
  doc.createParagraph('Butter').setNumbering(bulletListId, 0);
  doc.createParagraph('Cheese').setNumbering(bulletListId, 0);

  doc.createParagraph();

  // Another bullet list with different bullets
  doc.createParagraph('Project Tasks').setStyle('Heading2');

  // Create another bullet list with custom bullets
  const taskListId = doc.createBulletList(3, ['▪', '○', '▸']);

  doc.createParagraph('Complete Phase 1').setNumbering(taskListId, 0);
  doc.createParagraph('Complete Phase 2').setNumbering(taskListId, 0);
  doc.createParagraph('Complete Phase 3').setNumbering(taskListId, 0);
  doc.createParagraph('Begin Phase 4').setNumbering(taskListId, 0);

  doc.createParagraph();

  // Nested bullet list
  doc.createParagraph('Nested Bullet List').setStyle('Heading2');
  doc
    .createParagraph(
      'Bullet lists can have multiple levels for hierarchical information:'
    )
    .setStyle('Normal');
  doc.createParagraph();

  const nestedListId = doc.createBulletList();

  doc.createParagraph('Fruits').setNumbering(nestedListId, 0);
  doc.createParagraph('Apples').setNumbering(nestedListId, 1);
  doc.createParagraph('Oranges').setNumbering(nestedListId, 1);
  doc.createParagraph('Bananas').setNumbering(nestedListId, 1);

  doc.createParagraph('Vegetables').setNumbering(nestedListId, 0);
  doc.createParagraph('Carrots').setNumbering(nestedListId, 1);
  doc.createParagraph('Broccoli').setNumbering(nestedListId, 1);
  doc.createParagraph('Spinach').setNumbering(nestedListId, 1);

  doc.createParagraph('Dairy').setNumbering(nestedListId, 0);
  doc.createParagraph('Milk').setNumbering(nestedListId, 1);
  doc.createParagraph('Cheese').setNumbering(nestedListId, 1);
  doc.createParagraph('Yogurt').setNumbering(nestedListId, 1);

  doc.createParagraph();

  // Three-level nested list
  doc.createParagraph('Three-Level List').setStyle('Heading2');

  const deepListId = doc.createBulletList();

  doc.createParagraph('Programming Languages').setNumbering(deepListId, 0);
  doc.createParagraph('JavaScript').setNumbering(deepListId, 1);
  doc.createParagraph('Node.js runtime').setNumbering(deepListId, 2);
  doc.createParagraph('Browser runtime').setNumbering(deepListId, 2);
  doc.createParagraph('TypeScript').setNumbering(deepListId, 1);
  doc.createParagraph('Type safety').setNumbering(deepListId, 2);
  doc.createParagraph('Better IDE support').setNumbering(deepListId, 2);

  doc.createParagraph('Python').setNumbering(deepListId, 0);
  doc.createParagraph('Django framework').setNumbering(deepListId, 1);
  doc.createParagraph('Flask framework').setNumbering(deepListId, 1);

  // Summary
  doc.createParagraph();
  doc.createParagraph('Summary').setStyle('Heading1');
  doc
    .createParagraph(
      'This example demonstrated creating bullet lists with DocXML using the ' +
        'createBulletList() method and setNumbering() on paragraphs.'
    )
    .setStyle('Normal');

  // Save
  const filename = 'simple-bullet-list.docx';
  await doc.save(filename);
  console.log(`✓ Created ${filename}`);
  console.log('  Open in Microsoft Word to see bullet lists!');
}

// Run the example
demonstrateBulletList().catch(console.error);
