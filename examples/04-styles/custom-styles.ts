/**
 * Example: Creating Custom Styles
 *
 * Demonstrates how to create custom paragraph and character styles
 * with specific formatting, colors, and properties.
 */

import { Document, Style } from '../../src';

async function demonstrateCustomStyles() {
  console.log('Creating document with custom styles...\n');

  const doc = Document.create({
    properties: {
      title: 'Custom Styles Example',
      creator: 'DocXML',
    },
  });

  // 1. Create a custom "Alert" style for warnings
  console.log('Creating Alert style...');
  const alertStyle = Style.create({
    styleId: 'Alert',
    name: 'Alert',
    type: 'paragraph',
    basedOn: 'Normal',
    customStyle: true,
    paragraphFormatting: {
      alignment: 'center',
      spacing: {
        before: 240,
        after: 240,
      },
    },
    runFormatting: {
      bold: true,
      color: 'FF0000', // Red
      size: 12,
    },
  });
  doc.addStyle(alertStyle);

  // 2. Create a custom "CodeBlock" style for code snippets
  console.log('Creating CodeBlock style...');
  const codeBlockStyle = Style.create({
    styleId: 'CodeBlock',
    name: 'Code Block',
    type: 'paragraph',
    basedOn: 'Normal',
    customStyle: true,
    paragraphFormatting: {
      alignment: 'left',
      indentation: {
        left: 720, // 0.5 inch indent
      },
      spacing: {
        before: 120,
        after: 120,
      },
    },
    runFormatting: {
      font: 'Consolas',
      size: 10,
      color: '1F1F1F', // Dark gray
    },
  });
  doc.addStyle(codeBlockStyle);

  // 3. Create a custom "Highlight" style for important text
  console.log('Creating Highlight style...');
  const highlightStyle = Style.create({
    styleId: 'Highlight',
    name: 'Highlight',
    type: 'paragraph',
    basedOn: 'Normal',
    customStyle: true,
    paragraphFormatting: {
      alignment: 'justify',
      spacing: {
        before: 120,
        after: 120,
      },
    },
    runFormatting: {
      bold: true,
      color: '0066CC', // Blue
      highlight: 'yellow',
      size: 11,
    },
  });
  doc.addStyle(highlightStyle);

  // 4. Create a custom "Quote" style for quotations
  console.log('Creating Quote style...');
  const quoteStyle = Style.create({
    styleId: 'Quote',
    name: 'Quote',
    type: 'paragraph',
    basedOn: 'Normal',
    customStyle: true,
    paragraphFormatting: {
      alignment: 'justify',
      indentation: {
        left: 720,  // 0.5 inch
        right: 720, // 0.5 inch
      },
      spacing: {
        before: 240,
        after: 240,
      },
    },
    runFormatting: {
      italic: true,
      color: '595959', // Gray
      size: 11,
    },
  });
  doc.addStyle(quoteStyle);

  // 5. Create a custom "SectionTitle" style
  console.log('Creating SectionTitle style...');
  const sectionTitleStyle = Style.create({
    styleId: 'SectionTitle',
    name: 'Section Title',
    type: 'paragraph',
    basedOn: 'Heading2',
    customStyle: true,
    paragraphFormatting: {
      alignment: 'left',
      spacing: {
        before: 360,
        after: 180,
      },
      keepNext: true,
    },
    runFormatting: {
      font: 'Arial',
      size: 14,
      bold: true,
      color: '2E5C8A', // Custom blue
      allCaps: true,
    },
  });
  doc.addStyle(sectionTitleStyle);

  // Now use all the custom styles in a document
  doc.createParagraph('Custom Styles Demonstration').setStyle('Title');
  doc.createParagraph('Examples of custom paragraph styles').setStyle('Subtitle');
  doc.createParagraph();

  // Use Alert style
  doc.createParagraph('⚠ IMPORTANT: This is an alert message').setStyle('Alert');
  doc.createParagraph();

  // Regular content
  doc.createParagraph('Introduction').setStyle('Heading1');
  doc
    .createParagraph(
      'This document demonstrates various custom styles created with DocXML. ' +
        'Each style has unique formatting properties including fonts, colors, ' +
        'alignment, spacing, and more.'
    )
    .setStyle('Normal');
  doc.createParagraph();

  // Use SectionTitle style
  doc.createParagraph('Alert Style').setStyle('SectionTitle');
  doc
    .createParagraph(
      'The Alert style is perfect for warning messages. It features centered alignment, ' +
        'bold red text at 12pt, with extra spacing above and below to make it stand out.'
    )
    .setStyle('Normal');
  doc.createParagraph('⚠ WARNING: Do not proceed without reading this!').setStyle('Alert');
  doc.createParagraph();

  // Use CodeBlock style
  doc.createParagraph('Code Block Style').setStyle('SectionTitle');
  doc
    .createParagraph(
      'The CodeBlock style uses a monospace font (Consolas) and is perfect for ' +
        'displaying code snippets or technical content:'
    )
    .setStyle('Normal');
  doc.createParagraph('const doc = Document.create();').setStyle('CodeBlock');
  doc.createParagraph('doc.createParagraph("Hello World");').setStyle('CodeBlock');
  doc.createParagraph('await doc.save("output.docx");').setStyle('CodeBlock');
  doc.createParagraph();

  // Use Highlight style
  doc.createParagraph('Highlight Style').setStyle('SectionTitle');
  doc
    .createParagraph(
      'Regular text can be emphasized using the Highlight style, which combines ' +
        'bold blue text with yellow highlighting.'
    )
    .setStyle('Normal');
  doc
    .createParagraph(
      'This text is highlighted to draw attention to important information that ' +
        'readers should not miss. The yellow background combined with blue text ' +
        'creates excellent contrast.'
    )
    .setStyle('Highlight');
  doc.createParagraph();

  // Use Quote style
  doc.createParagraph('Quote Style').setStyle('SectionTitle');
  doc
    .createParagraph(
      'The Quote style is ideal for quotations, featuring italic text with indentation ' +
        'on both sides and extra spacing:'
    )
    .setStyle('Normal');
  doc
    .createParagraph(
      'The best way to predict the future is to invent it. This quote style makes ' +
        'quotations stand out with italics, gray color, and symmetric indentation ' +
        'that clearly distinguishes quoted material from regular text.'
    )
    .setStyle('Quote');
  doc.createParagraph();

  // Comparison
  doc.createParagraph('Style Comparison').setStyle('Heading1');
  doc.createParagraph('All Custom Styles').setStyle('SectionTitle');
  doc.createParagraph('Here are all the custom styles side by side:').setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('1. Alert Style Example').setStyle('Alert');
  doc.createParagraph('2. Code Block: function example() { return true; }').setStyle('CodeBlock');
  doc.createParagraph('3. Highlight Style Example - Important Information').setStyle('Highlight');
  doc.createParagraph('4. Quote Style - "Example quotation text"').setStyle('Quote');
  doc.createParagraph('5. Section Title Style').setStyle('SectionTitle');
  doc.createParagraph();

  // Technical details
  doc.createParagraph('Creating Your Own Styles').setStyle('Heading1');
  doc
    .createParagraph(
      'To create custom styles, use the Style.create() method with formatting properties:'
    )
    .setStyle('Normal');
  doc.createParagraph();

  doc.createParagraph('const style = Style.create({').setStyle('CodeBlock');
  doc.createParagraph('  styleId: "MyStyle",').setStyle('CodeBlock');
  doc.createParagraph('  name: "My Style",').setStyle('CodeBlock');
  doc.createParagraph('  type: "paragraph",').setStyle('CodeBlock');
  doc.createParagraph('  basedOn: "Normal",').setStyle('CodeBlock');
  doc.createParagraph('  runFormatting: { bold: true, color: "FF0000" }').setStyle('CodeBlock');
  doc.createParagraph('});').setStyle('CodeBlock');
  doc.createParagraph('doc.addStyle(style);').setStyle('CodeBlock');
  doc.createParagraph();

  doc.createParagraph('⚠ Remember to add the style before using it!').setStyle('Alert');

  // Save
  const filename = 'custom-styles.docx';
  await doc.save(filename);
  console.log(`\n✓ Created ${filename}`);
  console.log('  Open in Microsoft Word to see custom styles in action!');
  console.log('\nCustom styles created:');
  console.log('  • Alert (red, bold, centered)');
  console.log('  • CodeBlock (monospace, indented)');
  console.log('  • Highlight (blue, bold, yellow background)');
  console.log('  • Quote (italic, gray, indented)');
  console.log('  • SectionTitle (all caps, custom color)');
}

// Run the example
demonstrateCustomStyles().catch(console.error);
