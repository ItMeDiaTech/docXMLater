/**
 * Example: Simple Numbered List
 *
 * Demonstrates how to create numbered lists with decimal, roman numeral,
 * and alphabetic formats.
 */

import { Document } from '../../src';

async function demonstrateNumberedList() {
  console.log('Creating document with numbered lists...\n');

  const doc = Document.create({
    properties: {
      title: 'Numbered List Example',
      creator: 'DocXML',
    },
  });

  // Add title
  doc.createParagraph('Numbered List Example').setStyle('Title');
  doc.createParagraph();

  // Simple decimal list
  doc.createParagraph('Simple Decimal List').setStyle('Heading1');
  doc
    .createParagraph(
      'A numbered list uses sequential numbers for each item:'
    )
    .setStyle('Normal');
  doc.createParagraph();

  // Create decimal numbered list
  const decimalListId = doc.createNumberedList();

  doc.createParagraph('First step of the process').setNumbering(decimalListId, 0);
  doc.createParagraph('Second step of the process').setNumbering(decimalListId, 0);
  doc.createParagraph('Third step of the process').setNumbering(decimalListId, 0);
  doc.createParagraph('Fourth step of the process').setNumbering(decimalListId, 0);
  doc.createParagraph('Fifth step of the process').setNumbering(decimalListId, 0);

  doc.createParagraph();

  // Nested numbered list
  doc.createParagraph('Nested Numbered List').setStyle('Heading1');
  doc
    .createParagraph(
      'Numbered lists can have multiple levels with different formats:'
    )
    .setStyle('Normal');
  doc.createParagraph();

  // Create nested list with different formats per level
  const nestedListId = doc.createNumberedList(3, ['decimal', 'lowerLetter', 'lowerRoman']);

  doc.createParagraph('Main Topic One').setNumbering(nestedListId, 0);
  doc.createParagraph('Subtopic a').setNumbering(nestedListId, 1);
  doc.createParagraph('Detail i').setNumbering(nestedListId, 2);
  doc.createParagraph('Detail ii').setNumbering(nestedListId, 2);
  doc.createParagraph('Subtopic b').setNumbering(nestedListId, 1);
  doc.createParagraph('Detail i').setNumbering(nestedListId, 2);
  doc.createParagraph('Detail ii').setNumbering(nestedListId, 2);

  doc.createParagraph('Main Topic Two').setNumbering(nestedListId, 0);
  doc.createParagraph('Subtopic a').setNumbering(nestedListId, 1);
  doc.createParagraph('Detail i').setNumbering(nestedListId, 2);
  doc.createParagraph('Subtopic b').setNumbering(nestedListId, 1);

  doc.createParagraph();

  // Recipe example
  doc.createParagraph('Recipe Instructions').setStyle('Heading1');
  doc
    .createParagraph(
      'Here is a recipe using numbered steps:'
    )
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('Chocolate Chip Cookies').setStyle('Heading2');

  const recipeListId = doc.createNumberedList();

  doc.createParagraph('Preheat oven to 375°F (190°C)').setNumbering(recipeListId, 0);
  doc
    .createParagraph('Mix butter and sugars until creamy')
    .setNumbering(recipeListId, 0);
  doc
    .createParagraph('Beat in eggs and vanilla extract')
    .setNumbering(recipeListId, 0);
  doc
    .createParagraph('In separate bowl, combine flour, baking soda, and salt')
    .setNumbering(recipeListId, 0);
  doc
    .createParagraph('Gradually add dry ingredients to wet ingredients')
    .setNumbering(recipeListId, 0);
  doc.createParagraph('Stir in chocolate chips').setNumbering(recipeListId, 0);
  doc
    .createParagraph('Drop rounded spoonfuls onto ungreased cookie sheet')
    .setNumbering(recipeListId, 0);
  doc.createParagraph('Bake for 10-12 minutes').setNumbering(recipeListId, 0);
  doc.createParagraph('Cool on baking sheet for 2 minutes').setNumbering(recipeListId, 0);
  doc.createParagraph('Transfer to wire rack to cool completely').setNumbering(recipeListId, 0);

  doc.createParagraph();

  // Legal-style outline
  doc.createParagraph('Legal-Style Outline').setStyle('Heading1');
  doc
    .createParagraph(
      'Multi-level numbered lists are commonly used in legal documents and formal outlines:'
    )
    .setStyle('Normal');
  doc.createParagraph();

  const outlineListId = doc.createNumberedList(4, [
    'decimal',
    'lowerLetter',
    'lowerRoman',
    'decimal',
  ]);

  doc.createParagraph('Introduction').setNumbering(outlineListId, 0);
  doc.createParagraph('Background information').setNumbering(outlineListId, 1);
  doc.createParagraph('Historical context').setNumbering(outlineListId, 2);
  doc.createParagraph('Recent developments').setNumbering(outlineListId, 2);
  doc.createParagraph('Purpose of document').setNumbering(outlineListId, 1);

  doc.createParagraph('Main Argument').setNumbering(outlineListId, 0);
  doc.createParagraph('First point').setNumbering(outlineListId, 1);
  doc.createParagraph('Supporting evidence').setNumbering(outlineListId, 2);
  doc.createParagraph('Expert testimony').setNumbering(outlineListId, 3);
  doc.createParagraph('Statistical data').setNumbering(outlineListId, 3);
  doc.createParagraph('Counter-argument').setNumbering(outlineListId, 2);
  doc.createParagraph('Second point').setNumbering(outlineListId, 1);
  doc.createParagraph('Case studies').setNumbering(outlineListId, 2);

  doc.createParagraph('Conclusion').setNumbering(outlineListId, 0);
  doc.createParagraph('Summary of findings').setNumbering(outlineListId, 1);
  doc.createParagraph('Recommendations').setNumbering(outlineListId, 1);

  doc.createParagraph();

  // Mixed list with regular text
  doc.createParagraph('List with Interruptions').setStyle('Heading1');
  doc
    .createParagraph(
      'You can interrupt a list with regular paragraphs and continue numbering:'
    )
    .setStyle('Normal');
  doc.createParagraph();

  const mixedListId = doc.createNumberedList();

  doc.createParagraph('First numbered item').setNumbering(mixedListId, 0);
  doc.createParagraph('Second numbered item').setNumbering(mixedListId, 0);

  doc
    .createParagraph('This is a regular paragraph, not part of the list.')
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('Third numbered item (continues from 2)').setNumbering(mixedListId, 0);
  doc.createParagraph('Fourth numbered item').setNumbering(mixedListId, 0);

  // Summary
  doc.createParagraph();
  doc.createParagraph('Summary').setStyle('Heading1');
  doc
    .createParagraph(
      'This example demonstrated various numbered list formats including ' +
        'decimal (1, 2, 3), lower letters (a, b, c), and lower roman numerals (i, ii, iii). ' +
        'You can create nested lists with different formats at each level using ' +
        'createNumberedList() and setNumbering().'
    )
    .setStyle('Normal');

  // Save
  const filename = 'simple-numbered-list.docx';
  await doc.save(filename);
  console.log(`✓ Created ${filename}`);
  console.log('  Open in Microsoft Word to see numbered lists!');
}

// Run the example
demonstrateNumberedList().catch(console.error);
