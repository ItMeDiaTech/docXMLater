/**
 * Detached Paragraphs Example
 *
 * Demonstrates how to create paragraphs independently before adding them to a document.
 * This approach is useful for building reusable components, templates, or complex layouts.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import * as path from 'path';

async function main() {
  console.log('Creating document with detached paragraphs...');

  const doc = Document.create();

  // ============================================================================
  // 1. Basic Detached Paragraphs
  // ============================================================================

  // Create empty paragraph
  const emptyPara = Paragraph.create();

  // Create paragraph with text
  const textPara = Paragraph.create('This paragraph was created before being added to the document.');

  // Create paragraph with text and formatting
  const formattedPara = Paragraph.create('Centered paragraph with spacing', {
    alignment: 'center',
    spacing: { before: 240, after: 240 }
  });

  // Add to document
  doc.addParagraph(emptyPara);
  doc.addParagraph(textPara);
  doc.addParagraph(formattedPara);

  // ============================================================================
  // 2. Styled Detached Paragraphs
  // ============================================================================

  const heading1 = Paragraph.createWithStyle('Chapter 1: Introduction', 'Heading1');
  const heading2 = Paragraph.createWithStyle('Background', 'Heading2');

  doc.addParagraph(heading1);
  doc.addParagraph(heading2);

  // ============================================================================
  // 3. Complex Formatted Paragraphs
  // ============================================================================

  // Create paragraph with multiple formatted runs
  const complexPara = Paragraph.create()
    .addText('This paragraph has ', {})
    .addText('bold', { bold: true })
    .addText(', ', {})
    .addText('italic', { italic: true })
    .addText(', and ', {})
    .addText('colored', { color: 'FF0000' })
    .addText(' text.')
    .setAlignment('justify');

  doc.addParagraph(complexPara);

  // ============================================================================
  // 4. Using Paragraph.createFormatted()
  // ============================================================================

  const importantPara = Paragraph.createFormatted(
    'Important Notice',
    { bold: true, color: 'FF0000', size: 24 },
    { alignment: 'center', spacing: { before: 240, after: 240 } }
  );

  doc.addParagraph(importantPara);

  // ============================================================================
  // 5. Building Reusable Components
  // ============================================================================

  // Function to create a custom quote block
  function createQuoteBlock(text: string, author: string): Paragraph[] {
    const quotePara = Paragraph.create(text, {
      alignment: 'center',
      indentation: { left: 1440, right: 1440 },
      spacing: { before: 240, after: 120 }
    });
    quotePara.addText('', { italic: true });

    const authorPara = Paragraph.create(`— ${author}`, {
      alignment: 'right',
      indentation: { right: 1440 },
      spacing: { after: 240 }
    });

    return [quotePara, authorPara];
  }

  // Use the component
  const quote = createQuoteBlock(
    'The only way to do great work is to love what you do.',
    'Steve Jobs'
  );

  doc.addParagraph(Paragraph.createEmpty()); // Blank line
  quote.forEach(p => doc.addParagraph(p));

  // ============================================================================
  // 6. Paragraph Templates
  // ============================================================================

  // Create paragraph templates that can be cloned and reused
  const warningTemplate = Paragraph.createFormatted(
    'WARNING: ',
    { bold: true, color: 'FF6600' },
    {
      alignment: 'left',
      spacing: { before: 120, after: 120 },
      indentation: { left: 240 }
    }
  );

  // Clone and customize
  const warning1 = warningTemplate.clone();
  warning1.addText('Please read the documentation before proceeding.');

  const warning2 = warningTemplate.clone();
  warning2.addText('This operation cannot be undone.');

  doc.addParagraph(Paragraph.createEmpty());
  doc.addParagraph(warning1);
  doc.addParagraph(warning2);

  // ============================================================================
  // 7. Conditional Paragraph Building
  // ============================================================================

  const userData = {
    name: 'John Doe',
    isPremium: true,
    points: 1250
  };

  // Build paragraph based on data
  const greetingPara = Paragraph.create();
  greetingPara.addText('Hello, ', {});
  greetingPara.addText(userData.name, { bold: true });
  greetingPara.addText('!', {});

  if (userData.isPremium) {
    greetingPara.addText(' ', {});
    greetingPara.addText('(Premium Member)', { color: 'FFD700', italic: true });
  }

  const pointsPara = Paragraph.create(`You have ${userData.points} points.`, {
    indentation: { left: 240 }
  });

  doc.addParagraph(Paragraph.createEmpty());
  doc.addParagraph(greetingPara);
  doc.addParagraph(pointsPara);

  // ============================================================================
  // 8. Array-Based Document Building
  // ============================================================================

  const items = [
    { title: 'First Item', description: 'Description of first item' },
    { title: 'Second Item', description: 'Description of second item' },
    { title: 'Third Item', description: 'Description of third item' }
  ];

  doc.addParagraph(Paragraph.createEmpty());
  doc.addParagraph(Paragraph.createWithStyle('Items List', 'Heading2'));

  items.forEach((item, index) => {
    const titlePara = Paragraph.create(`${index + 1}. `, { spacing: { before: 120 } });
    titlePara.addText(item.title, { bold: true });

    const descPara = Paragraph.create(item.description, {
      indentation: { left: 360 },
      spacing: { after: 120 }
    });

    doc.addParagraph(titlePara);
    doc.addParagraph(descPara);
  });

  // ============================================================================
  // 9. Pre-building Content Before Conditions
  // ============================================================================

  // Build all possible paragraphs
  const successPara = Paragraph.create('Operation completed successfully!', {
    alignment: 'center'
  });
  successPara.addText('', { bold: true, color: '00AA00' });

  const errorPara = Paragraph.create('An error occurred. Please try again.', {
    alignment: 'center'
  });
  errorPara.addText('', { bold: true, color: 'FF0000' });

  // Conditionally add based on result
  const operationSuccess = true;
  doc.addParagraph(Paragraph.createEmpty());
  doc.addParagraph(operationSuccess ? successPara : errorPara);

  // ============================================================================
  // 10. Complex Formatting Example
  // ============================================================================

  const complexFormatted = Paragraph.create()
    .addText('Complex Formatting Example')
    .setAlignment('center')
    .setLeftIndent(720)
    .setRightIndent(720)
    .setSpaceBefore(240)
    .setSpaceAfter(240)
    .setLineSpacing(360, 'exact')
    .setKeepNext(true)
    .setKeepLines(true);

  doc.addParagraph(Paragraph.createEmpty());
  doc.addParagraph(complexFormatted);

  // Save document
  const outputPath = path.join(__dirname, '..', 'output', 'detached-paragraphs.docx');
  await doc.save(outputPath);

  console.log(`Document saved to: ${outputPath}`);
  console.log('\nKey Benefits of Detached Paragraphs:');
  console.log('1. Build content separately from document structure');
  console.log('2. Create reusable paragraph components and templates');
  console.log('3. Clone and customize existing paragraphs');
  console.log('4. Conditionally add paragraphs based on logic');
  console.log('5. Build content from arrays and data structures');
  console.log('6. Better separation of concerns in code');
}

main().catch(console.error);
