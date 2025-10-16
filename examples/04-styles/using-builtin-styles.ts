/**
 * Example: Using Built-in Styles
 *
 * Demonstrates all 13 built-in styles that come with every DocXML document:
 * - Normal (default paragraph)
 * - Heading1 through Heading9
 * - Title
 * - Subtitle
 * - ListParagraph
 */

import { Document } from '../../src';

async function demonstrateBuiltInStyles() {
  console.log('Creating document with all built-in styles...\n');

  const doc = Document.create({
    properties: {
      title: 'Built-in Styles Demonstration',
      creator: 'DocXML',
      subject: 'All 13 Built-in Styles',
    },
  });

  // Title and Subtitle styles
  doc.createParagraph('Built-in Styles Reference').setStyle('Title');
  doc
    .createParagraph('A comprehensive demonstration of all 13 built-in styles')
    .setStyle('Subtitle');
  doc.createParagraph();

  // Normal style (default body text)
  doc
    .createParagraph(
      'The Normal style is the default paragraph style used for body text. ' +
        'It uses Calibri 11pt font with standard spacing.'
    )
    .setStyle('Normal');
  doc.createParagraph();

  // Heading styles (9 levels)
  doc.createParagraph('Chapter 1: Heading Styles').setStyle('Heading1');
  doc
    .createParagraph(
      'Heading1 is the largest heading style, typically used for chapter titles. ' +
        'It uses Calibri Light 16pt, bold, in blue color.'
    )
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('Section 1.1: Heading Level 2').setStyle('Heading2');
  doc
    .createParagraph(
      'Heading2 is used for major sections within a chapter. ' +
        'It uses Calibri Light 13pt, bold, in a darker blue.'
    )
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('Subsection 1.1.1: Heading Level 3').setStyle('Heading3');
  doc
    .createParagraph(
      'Heading3 is for subsections. It uses Calibri Light 12pt, bold.'
    )
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('Minor Heading: Level 4').setStyle('Heading4');
  doc
    .createParagraph('Heading4 uses Calibri Light 11pt, bold, in blue.')
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('Lower Level: Heading 5').setStyle('Heading5');
  doc
    .createParagraph('Heading5 through 9 use Calibri Light 11pt (not bold).')
    .setStyle('Normal');
  doc.createParagraph();

  // Demonstrate remaining heading levels
  doc.createParagraph('Heading Level 6 Example').setStyle('Heading6');
  doc.createParagraph('Heading Level 7 Example').setStyle('Heading7');
  doc.createParagraph('Heading Level 8 Example').setStyle('Heading8');
  doc.createParagraph('Heading Level 9 Example').setStyle('Heading9');
  doc
    .createParagraph('All lower heading levels (5-9) maintain the same formatting.')
    .setStyle('Normal');
  doc.createParagraph();

  // Chapter 2: Special Styles
  doc.createParagraph('Chapter 2: Special Styles').setStyle('Heading1');

  doc.createParagraph('Title Style').setStyle('Heading2');
  doc
    .createParagraph(
      'The Title style is designed for document titles on cover pages. ' +
        'It uses Calibri Light 28pt in blue - the largest built-in style.'
    )
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('Subtitle Style').setStyle('Heading2');
  doc
    .createParagraph(
      'The Subtitle style complements the Title style for secondary information. ' +
        'It uses Calibri Light 14pt, italic, in gray color.'
    )
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('ListParagraph Style').setStyle('Heading2');
  doc
    .createParagraph(
      'The ListParagraph style is used for list items. It is based on Normal ' +
        'but includes a 0.5 inch left indent. Examples below:'
    )
    .setStyle('Normal');
  doc.createParagraph('First list item').setStyle('ListParagraph');
  doc.createParagraph('Second list item').setStyle('ListParagraph');
  doc.createParagraph('Third list item').setStyle('ListParagraph');
  doc.createParagraph();

  // Style comparison table (using paragraphs)
  doc.createParagraph('Chapter 3: Style Comparison').setStyle('Heading1');
  doc
    .createParagraph(
      'Here is a visual comparison of all paragraph styles in action:'
    )
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('This is Normal style').setStyle('Normal');
  doc.createParagraph('This is Heading1 style').setStyle('Heading1');
  doc.createParagraph('This is Heading2 style').setStyle('Heading2');
  doc.createParagraph('This is Heading3 style').setStyle('Heading3');
  doc.createParagraph('This is Heading4 style').setStyle('Heading4');
  doc.createParagraph('This is Heading5 style').setStyle('Heading5');
  doc.createParagraph('This is Title style').setStyle('Title');
  doc.createParagraph('This is Subtitle style').setStyle('Subtitle');
  doc.createParagraph('This is ListParagraph style').setStyle('ListParagraph');
  doc.createParagraph();

  // Technical details
  doc.createParagraph('Technical Details').setStyle('Heading1');
  doc
    .createParagraph(
      'All built-in styles are automatically available in every document created ' +
        'with Document.create(). They follow Microsoft Word conventions and are ' +
        'fully compatible with Word 2016+.'
    )
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('Key Features:').setStyle('Heading2');
  doc
    .createParagraph(
      '• All headings use "Keep with Next" to stay with following content'
    )
    .setStyle('Normal');
  doc
    .createParagraph('• All headings use "Keep Lines" to prevent page breaks within')
    .setStyle('Normal');
  doc
    .createParagraph('• Consistent spacing and alignment throughout')
    .setStyle('Normal');
  doc.createParagraph('• Professional typography with Calibri font family').setStyle('Normal');
  doc
    .createParagraph('• Hierarchical color scheme using blues and grays')
    .setStyle('Normal');

  // Save
  const filename = 'using-builtin-styles.docx';
  await doc.save(filename);
  console.log(`✓ Created ${filename}`);
  console.log('  Open in Microsoft Word to see all 13 built-in styles in action!');
  console.log('\nStyles demonstrated:');
  console.log('  1. Normal');
  console.log('  2-10. Heading1 through Heading9');
  console.log('  11. Title');
  console.log('  12. Subtitle');
  console.log('  13. ListParagraph');
}

// Run the example
demonstrateBuiltInStyles().catch(console.error);
