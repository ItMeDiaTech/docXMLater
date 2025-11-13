/**
 * Example: Style Inheritance
 *
 * Demonstrates how styles inherit from parent styles using the basedOn property,
 * creating a hierarchy of styles that share common formatting while allowing
 * customization at each level.
 */

import { Document, Style } from '../../src';

async function demonstrateStyleInheritance() {
  console.log('Creating document demonstrating style inheritance...\n');

  const doc = Document.create({
    properties: {
      title: 'Style Inheritance Example',
      creator: 'DocXML',
    },
  });

  // Create a base style hierarchy
  console.log('Building style hierarchy...');

  // 1. Base Document Style - The root of our custom hierarchy
  const baseDocStyle = Style.create({
    styleId: 'BaseDoc',
    name: 'Base Document',
    type: 'paragraph',
    basedOn: 'Normal', // Based on built-in Normal
    customStyle: true,
    runFormatting: {
      font: 'Georgia', // Serif font for the whole document
      size: 11,
      color: '1A1A1A', // Almost black
    },
    paragraphFormatting: {
      alignment: 'justify',
      spacing: {
        after: 160,
        line: 276,
      },
    },
  });
  doc.addStyle(baseDocStyle);

  // 2. Base Heading Style - Parent for all custom headings
  const baseHeadingStyle = Style.create({
    styleId: 'BaseHeading',
    name: 'Base Heading',
    type: 'paragraph',
    basedOn: 'BaseDoc', // Inherits from BaseDoc (font, spacing)
    customStyle: true,
    runFormatting: {
      font: 'Arial', // Override: Sans-serif for headings
      bold: true,
      color: '0B3D91', // NASA blue
    },
    paragraphFormatting: {
      keepNext: true,
      keepLines: true,
      spacing: {
        before: 240,
        after: 120,
      },
    },
  });
  doc.addStyle(baseHeadingStyle);

  // 3. Main Heading - Level 1 (inherits from BaseHeading)
  const mainHeadingStyle = Style.create({
    styleId: 'MainHeading',
    name: 'Main Heading',
    type: 'paragraph',
    basedOn: 'BaseHeading', // Inherits bold, Arial, blue, keepNext
    customStyle: true,
    runFormatting: {
      size: 18, // Override: Larger size
      allCaps: true, // Add: All caps
    },
    paragraphFormatting: {
      spacing: {
        before: 360, // Override: More space before
        after: 180,
      },
    },
  });
  doc.addStyle(mainHeadingStyle);

  // 4. Sub Heading - Level 2 (inherits from BaseHeading)
  const subHeadingStyle = Style.create({
    styleId: 'SubHeading',
    name: 'Sub Heading',
    type: 'paragraph',
    basedOn: 'BaseHeading', // Same parent as MainHeading
    customStyle: true,
    runFormatting: {
      size: 14, // Override: Medium size
      italic: true, // Add: Italic
    },
    paragraphFormatting: {
      indentation: {
        left: 360, // Add: Slight indent
      },
    },
  });
  doc.addStyle(subHeadingStyle);

  // 5. Emphasis Text - Special body text (inherits from BaseDoc)
  const emphasisStyle = Style.create({
    styleId: 'Emphasis',
    name: 'Emphasis',
    type: 'paragraph',
    basedOn: 'BaseDoc', // Same font and base formatting
    customStyle: true,
    runFormatting: {
      bold: true, // Add: Bold
      color: '8B0000', // Override: Dark red instead of black
    },
    paragraphFormatting: {
      indentation: {
        left: 720,
      },
      spacing: {
        before: 120,
        after: 120,
      },
    },
  });
  doc.addStyle(emphasisStyle);

  // 6. Fine Print - Small text (inherits from BaseDoc)
  const finePrintStyle = Style.create({
    styleId: 'FinePrint',
    name: 'Fine Print',
    type: 'paragraph',
    basedOn: 'BaseDoc',
    customStyle: true,
    runFormatting: {
      size: 9, // Override: Smaller
      color: '666666', // Override: Gray
      italic: true, // Add: Italic
    },
    paragraphFormatting: {
      alignment: 'left', // Override: Left instead of justify
      spacing: {
        after: 80, // Override: Less space
      },
    },
  });
  doc.addStyle(finePrintStyle);

  // Now demonstrate the hierarchy in the document
  doc.createParagraph('Style Inheritance in DocXML').setStyle('Title');
  doc.createParagraph('Understanding the basedOn property').setStyle('Subtitle');
  doc.createParagraph();

  // Show the hierarchy
  doc.createParagraph('Understanding Style Hierarchy').setStyle('MainHeading');
  doc
    .createParagraph(
      'Style inheritance allows you to create a hierarchy of styles where child styles ' +
        'inherit properties from parent styles. This creates consistency while allowing ' +
        'customization at each level.'
    )
    .setStyle('BaseDoc');
  doc.createParagraph();

  doc.createParagraph('Our Custom Style Hierarchy').setStyle('SubHeading');
  doc
    .createParagraph(
      'This document uses a custom style hierarchy built on top of the Normal style:'
    )
    .setStyle('BaseDoc');
  doc.createParagraph();

  doc.createParagraph('Normal (built-in)').setStyle('FinePrint');
  doc.createParagraph('  └─ BaseDoc (Georgia, justified)').setStyle('FinePrint');
  doc.createParagraph('      ├─ BaseHeading (Arial, bold, blue)').setStyle('FinePrint');
  doc.createParagraph('      │   ├─ MainHeading (18pt, all caps)').setStyle('FinePrint');
  doc.createParagraph('      │   └─ SubHeading (14pt, italic)').setStyle('FinePrint');
  doc.createParagraph('      ├─ Emphasis (bold, dark red)').setStyle('FinePrint');
  doc.createParagraph('      └─ FinePrint (9pt, gray, italic)').setStyle('FinePrint');
  doc.createParagraph();

  // Demonstrate each style
  doc.createParagraph('Demonstration of Each Style').setStyle('MainHeading');

  doc.createParagraph('BaseDoc Style').setStyle('SubHeading');
  doc
    .createParagraph(
      'This is the BaseDoc style. It uses Georgia font (serif) at 11pt with justified ' +
        'alignment. All body text in this document inherits from BaseDoc, creating a ' +
        'consistent typographic foundation. Notice the justified alignment and serif font.'
    )
    .setStyle('BaseDoc');
  doc.createParagraph();

  doc.createParagraph('Main Heading Style').setStyle('SubHeading');
  doc
    .createParagraph(
      'The MainHeading style (shown above) inherits from BaseHeading, which means it ' +
        'automatically gets Arial font, bold weight, and blue color. It adds 18pt size ' +
        'and all caps transformation.'
    )
    .setStyle('BaseDoc');
  doc.createParagraph();

  doc.createParagraph('Sub Heading Style').setStyle('SubHeading');
  doc
    .createParagraph(
      'The SubHeading style (also shown above) inherits the same BaseHeading properties ' +
        'but uses 14pt size and adds italic styling. It also includes a left indent to ' +
        'show hierarchy.'
    )
    .setStyle('BaseDoc');
  doc.createParagraph();

  doc.createParagraph('Emphasis Style').setStyle('SubHeading');
  doc
    .createParagraph(
      'The Emphasis style inherits from BaseDoc (not BaseHeading), so it maintains the ' +
        'Georgia font. It adds bold weight and uses dark red color to draw attention.'
    )
    .setStyle('Emphasis');
  doc.createParagraph();

  doc.createParagraph('Fine Print Style').setStyle('SubHeading');
  doc
    .createParagraph(
      'This is FinePrint style - small (9pt), gray, italic text for disclaimers and notes.'
    )
    .setStyle('FinePrint');
  doc.createParagraph();

  // Benefits section
  doc.createParagraph('Benefits of Style Inheritance').setStyle('MainHeading');

  doc.createParagraph('Benefit 1: Consistency').setStyle('SubHeading');
  doc
    .createParagraph(
      'All headings automatically share the same font (Arial), weight (bold), and color ' +
        '(blue) because they inherit from BaseHeading. Change BaseHeading and all children ' +
        'update automatically.'
    )
    .setStyle('BaseDoc');
  doc.createParagraph();

  doc.createParagraph('Benefit 2: Maintainability').setStyle('SubHeading');
  doc
    .createParagraph(
      'Want to change the document font from Georgia to another serif font? Just update ' +
        'BaseDoc and all body styles (Emphasis, FinePrint) inherit the change.'
    )
    .setStyle('BaseDoc');
  doc.createParagraph();

  doc.createParagraph('Benefit 3: Flexibility').setStyle('SubHeading');
  doc
    .createParagraph(
      'Each child style can override specific properties while keeping inherited ones. ' +
        'SubHeading overrides size and adds italic, but keeps the blue color from BaseHeading.'
    )
    .setStyle('BaseDoc');
  doc.createParagraph();

  // Code example
  doc.createParagraph('Creating Inherited Styles in Code').setStyle('MainHeading');
  doc
    .createParagraph(
      'Here is how you create a child style that inherits from a parent:'
    )
    .setStyle('BaseDoc');
  doc.createParagraph();

  doc.createParagraph('// Create parent style').setStyle('FinePrint');
  doc.createParagraph('const parentStyle = Style.create({').setStyle('FinePrint');
  doc.createParagraph('  styleId: "Parent",').setStyle('FinePrint');
  doc.createParagraph('  basedOn: "Normal",').setStyle('FinePrint');
  doc.createParagraph('  runFormatting: { font: "Arial", size: 11 }').setStyle('FinePrint');
  doc.createParagraph('});').setStyle('FinePrint');
  doc.createParagraph('doc.addStyle(parentStyle);').setStyle('FinePrint');
  doc.createParagraph();
  doc.createParagraph('// Create child style that inherits').setStyle('FinePrint');
  doc.createParagraph('const childStyle = Style.create({').setStyle('FinePrint');
  doc.createParagraph('  styleId: "Child",').setStyle('FinePrint');
  doc.createParagraph('  basedOn: "Parent", // <-- Inherits font and size').setStyle('FinePrint');
  doc.createParagraph('  runFormatting: { bold: true } // <-- Adds bold').setStyle('FinePrint');
  doc.createParagraph('});').setStyle('FinePrint');
  doc.createParagraph('doc.addStyle(childStyle);').setStyle('FinePrint');
  doc.createParagraph();

  // Key takeaway
  doc
    .createParagraph(
      'Key Takeaway: Use basedOn to create style hierarchies that promote consistency and ' +
        'maintainability in your documents. Start with a base style for your document, create ' +
        'specialized parents for headings and special content, then derive specific styles as needed.'
    )
    .setStyle('Emphasis');

  // Save
  const filename = 'style-inheritance.docx';
  await doc.save(filename);
  console.log(`✓ Created ${filename}`);
  console.log('  Open in Microsoft Word to see style inheritance in action!');
  console.log('\nStyle hierarchy created:');
  console.log('  Normal → BaseDoc → Emphasis');
  console.log('  Normal → BaseDoc → FinePrint');
  console.log('  Normal → BaseDoc → BaseHeading → MainHeading');
  console.log('  Normal → BaseDoc → BaseHeading → SubHeading');
  console.log('\nNotice how each child inherits and optionally overrides parent properties.');
}

// Run the example
demonstrateStyleInheritance().catch(console.error);
