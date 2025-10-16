/**
 * Examples showing the Document class - the high-level API
 */

import { Document } from '../src';

/**
 * Example 1: Create a simple document
 */
async function example1SimpleDocument() {
  console.log('\n=== Example 1: Simple Document ===');

  const doc = Document.create();

  doc.createParagraph('Hello, World!');
  doc.createParagraph('This is a simple document created with the Document class.');

  await doc.save('example1-simple-document.docx');
  console.log('✓ Created example1-simple-document.docx');
}

/**
 * Example 2: Document with properties
 */
async function example2DocumentProperties() {
  console.log('\n=== Example 2: Document with Properties ===');

  const doc = Document.create({
    properties: {
      title: 'My Document',
      subject: 'Document Example',
      creator: 'DocXML Framework',
      keywords: 'example, docx, document',
      description: 'A demonstration of document properties',
    },
  });

  doc.createParagraph('Document with metadata');
  doc.createParagraph('Check the properties in Word!');

  await doc.save('example2-document-properties.docx');
  console.log('✓ Created example2-document-properties.docx');

  // Display properties
  const props = doc.getProperties();
  console.log(`  Title: ${props.title}`);
  console.log(`  Creator: ${props.creator}`);
}

/**
 * Example 3: Formatted document
 */
async function example3FormattedDocument() {
  console.log('\n=== Example 3: Formatted Document ===');

  const doc = Document.create();

  // Title
  const title = doc.createParagraph();
  title.setAlignment('center');
  title.setSpaceBefore(480);
  title.setSpaceAfter(240);
  title.addText('Formatted Document Example', {
    bold: true,
    size: 18,
    color: '0066CC',
  });

  // Introduction
  const intro = doc.createParagraph();
  intro.addText('This document demonstrates various formatting options available in DocXML.');
  intro.setSpaceAfter(240);

  // Section heading
  const heading = doc.createParagraph();
  heading.addText('Text Formatting', { bold: true, size: 14 });
  heading.setSpaceAfter(120);

  // Content with mixed formatting
  const content = doc.createParagraph();
  content.addText('You can use ');
  content.addText('bold', { bold: true });
  content.addText(', ');
  content.addText('italic', { italic: true });
  content.addText(', ');
  content.addText('underlined', { underline: 'single' });
  content.addText(', and ');
  content.addText('colored', { color: 'FF0000' });
  content.addText(' text easily.');

  await doc.save('example3-formatted-document.docx');
  console.log('✓ Created example3-formatted-document.docx');
  console.log(`  Total paragraphs: ${doc.getParagraphCount()}`);
}

/**
 * Example 4: Multi-paragraph document
 */
async function example4MultiParagraph() {
  console.log('\n=== Example 4: Multi-Paragraph Document ===');

  const doc = Document.create({
    properties: {
      title: 'Multi-Paragraph Example',
      creator: 'DocXML',
    },
  });

  // Add title
  doc.createParagraph().setAlignment('center').addText('The DocXML Framework', {
    bold: true,
    size: 20,
  });

  doc.createParagraph(); // Empty line

  // Add multiple content paragraphs
  doc.createParagraph('DocXML is a comprehensive framework for creating and manipulating Microsoft Word documents programmatically.');

  doc.createParagraph('It provides a clean, intuitive API that makes it easy to generate professional documents with rich formatting.');

  const features = doc.createParagraph();
  features.addText('Key Features:', { bold: true });

  doc.createParagraph('• Simple, high-level API');
  doc.createParagraph('• Full TypeScript support');
  doc.createParagraph('• Comprehensive formatting options');
  doc.createParagraph('• Production-ready and well-tested');

  await doc.save('example4-multi-paragraph.docx');
  console.log('✓ Created example4-multi-paragraph.docx');
  console.log(`  Total paragraphs: ${doc.getParagraphCount()}`);
}

/**
 * Example 5: Loading and modifying documents
 */
async function example5LoadAndModify() {
  console.log('\n=== Example 5: Load and Modify ===');

  // Create initial document
  const doc1 = Document.create({
    properties: { title: 'Original Document' },
  });
  doc1.createParagraph('This is the original content.');
  await doc1.save('example5-original.docx');
  console.log('✓ Created example5-original.docx');

  // Load and modify
  const doc2 = await Document.load('example5-original.docx');

  doc2.setProperties({ title: 'Modified Document' });
  doc2.createParagraph('This paragraph was added after loading.');
  doc2.createParagraph().addText('Modified content!', { bold: true, color: 'FF0000' });

  await doc2.save('example5-modified.docx');
  console.log('✓ Created example5-modified.docx');
  console.log(`  Original paragraphs: 1`);
  console.log(`  After modification: ${doc2.getParagraphCount()}`);
}

/**
 * Example 6: Working with buffers
 */
async function example6Buffers() {
  console.log('\n=== Example 6: Working with Buffers ===');

  const doc1 = Document.create();
  doc1.createParagraph('This document was created and saved to a buffer.');

  // Save to buffer
  const buffer = await doc1.toBuffer();
  console.log(`✓ Generated buffer (${buffer.length} bytes)`);

  // Load from buffer
  const doc2 = await Document.loadFromBuffer(buffer);
  doc2.createParagraph('This was added after loading from buffer.');

  await doc2.save('example6-buffer.docx');
  console.log('✓ Created example6-buffer.docx from buffer');
}

// Run all examples
async function runExamples() {
  console.log('=== DocXML Document Class Examples ===');

  try {
    await example1SimpleDocument();
    await example2DocumentProperties();
    await example3FormattedDocument();
    await example4MultiParagraph();
    await example5LoadAndModify();
    await example6Buffers();

    console.log('\n=== All examples completed successfully! ===');
    console.log('\nGenerated files:');
    console.log('  - example1-simple-document.docx');
    console.log('  - example2-document-properties.docx');
    console.log('  - example3-formatted-document.docx');
    console.log('  - example4-multi-paragraph.docx');
    console.log('  - example5-original.docx');
    console.log('  - example5-modified.docx');
    console.log('  - example6-buffer.docx');
  } catch (error) {
    console.error('Error running examples:', error);
    process.exit(1);
  }
}

// Run if executed directly
if (require.main === module) {
  runExamples();
}
