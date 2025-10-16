/**
 * Examples showing Paragraph and Run usage
 */

import { ZipHandler, DOCX_PATHS, Paragraph, XMLBuilder } from '../src';

/**
 * Example 1: Create a simple paragraph with formatted text
 */
async function example1SimpleFormattedText() {
  console.log('\n=== Example 1: Simple Formatted Text ===');

  const handler = new ZipHandler();

  // Create a paragraph with mixed formatting
  const para = new Paragraph();
  para.addText('This is normal text. ');
  para.addText('This is bold text. ', { bold: true });
  para.addText('This is italic text. ', { italic: true });
  para.addText('This is both bold and italic!', { bold: true, italic: true });

  // Generate XML
  const paraXml = para.toXML();
  const docXml = XMLBuilder.createDocument([paraXml]);

  // Add to DOCX
  handler.addFile(DOCX_PATHS.CONTENT_TYPES, createContentTypes());
  handler.addFile(DOCX_PATHS.RELS, createRels());
  handler.addFile(DOCX_PATHS.DOCUMENT, docXml);

  await handler.save('example1-formatted-text.docx');
  console.log('✓ Created example1-formatted-text.docx');
  console.log(`  Text: "${para.getText()}"`);
}

/**
 * Example 2: Paragraph with alignment and spacing
 */
async function example2AlignmentAndSpacing() {
  console.log('\n=== Example 2: Alignment and Spacing ===');

  const handler = new ZipHandler();

  // Create paragraphs with different alignments
  const paragraphs: Paragraph[] = [];

  const leftPara = new Paragraph()
    .setAlignment('left')
    .setSpaceAfter(240)
    .addText('This paragraph is left-aligned.');
  paragraphs.push(leftPara);

  const centerPara = new Paragraph()
    .setAlignment('center')
    .setSpaceAfter(240)
    .addText('This paragraph is centered.', { bold: true });
  paragraphs.push(centerPara);

  const rightPara = new Paragraph()
    .setAlignment('right')
    .setSpaceAfter(240)
    .addText('This paragraph is right-aligned.', { italic: true });
  paragraphs.push(rightPara);

  const justifyPara = new Paragraph()
    .setAlignment('justify')
    .addText('This paragraph is justified. It has enough text to demonstrate justification across multiple lines when rendered in Word.');
  paragraphs.push(justifyPara);

  // Generate document
  const paraXmls = paragraphs.map(p => p.toXML());
  const docXml = XMLBuilder.createDocument(paraXmls);

  handler.addFile(DOCX_PATHS.CONTENT_TYPES, createContentTypes());
  handler.addFile(DOCX_PATHS.RELS, createRels());
  handler.addFile(DOCX_PATHS.DOCUMENT, docXml);

  await handler.save('example2-alignment-spacing.docx');
  console.log('✓ Created example2-alignment-spacing.docx');
}

/**
 * Example 3: Advanced text formatting
 */
async function example3AdvancedFormatting() {
  console.log('\n=== Example 3: Advanced Formatting ===');

  const handler = new ZipHandler();
  const paragraphs: Paragraph[] = [];

  // Font styles
  const fontPara = new Paragraph();
  fontPara.addText('Arial 12pt', { font: 'Arial', size: 12 });
  fontPara.addText(' | ');
  fontPara.addText('Times New Roman 14pt', { font: 'Times New Roman', size: 14 });
  fontPara.addText(' | ');
  fontPara.addText('Courier New 10pt', { font: 'Courier New', size: 10 });
  paragraphs.push(fontPara);

  // Colors
  const colorPara = new Paragraph().setSpaceBefore(240);
  colorPara.addText('Red ', { color: 'FF0000' });
  colorPara.addText('Green ', { color: '00FF00' });
  colorPara.addText('Blue ', { color: '0000FF' });
  colorPara.addText('Purple', { color: '800080' });
  paragraphs.push(colorPara);

  // Highlights
  const highlightPara = new Paragraph().setSpaceBefore(240);
  highlightPara.addText('Yellow highlight', { highlight: 'yellow' });
  highlightPara.addText(' ');
  highlightPara.addText('Green highlight', { highlight: 'green' });
  highlightPara.addText(' ');
  highlightPara.addText('Cyan highlight', { highlight: 'cyan' });
  paragraphs.push(highlightPara);

  // Underline styles
  const underlinePara = new Paragraph().setSpaceBefore(240);
  underlinePara.addText('Single underline', { underline: 'single' });
  underlinePara.addText(' | ');
  underlinePara.addText('Double underline', { underline: 'double' });
  underlinePara.addText(' | ');
  underlinePara.addText('Dotted underline', { underline: 'dotted' });
  paragraphs.push(underlinePara);

  // Subscript and superscript
  const scriptPara = new Paragraph().setSpaceBefore(240);
  scriptPara.addText('H');
  scriptPara.addText('2', { subscript: true });
  scriptPara.addText('O is water, and E=mc');
  scriptPara.addText('2', { superscript: true });
  scriptPara.addText(' is Einstein\'s equation.');
  paragraphs.push(scriptPara);

  // Text effects
  const effectsPara = new Paragraph().setSpaceBefore(240);
  effectsPara.addText('Strikethrough', { strike: true });
  effectsPara.addText(' | ');
  effectsPara.addText('Small Caps', { smallCaps: true });
  effectsPara.addText(' | ');
  effectsPara.addText('ALL CAPS', { allCaps: true });
  paragraphs.push(effectsPara);

  // Generate document
  const paraXmls = paragraphs.map(p => p.toXML());
  const docXml = XMLBuilder.createDocument(paraXmls);

  handler.addFile(DOCX_PATHS.CONTENT_TYPES, createContentTypes());
  handler.addFile(DOCX_PATHS.RELS, createRels());
  handler.addFile(DOCX_PATHS.DOCUMENT, docXml);

  await handler.save('example3-advanced-formatting.docx');
  console.log('✓ Created example3-advanced-formatting.docx');
}

/**
 * Example 4: Paragraph indentation
 */
async function example4Indentation() {
  console.log('\n=== Example 4: Indentation ===');

  const handler = new ZipHandler();
  const paragraphs: Paragraph[] = [];

  // No indentation
  paragraphs.push(
    new Paragraph()
      .addText('No indentation - this is the baseline paragraph.')
      .setSpaceAfter(120)
  );

  // Left indent
  paragraphs.push(
    new Paragraph()
      .setLeftIndent(720) // 0.5 inches
      .addText('Left indented by 0.5 inches.')
      .setSpaceAfter(120)
  );

  // First line indent
  paragraphs.push(
    new Paragraph()
      .setFirstLineIndent(720)
      .addText('This paragraph has a first line indent of 0.5 inches. The rest of the text wraps normally without indentation.')
      .setSpaceAfter(120)
  );

  // Both indents
  paragraphs.push(
    new Paragraph()
      .setLeftIndent(720)
      .setFirstLineIndent(720)
      .addText('This paragraph has both left indent and first line indent, creating a double indent effect.')
  );

  // Generate document
  const paraXmls = paragraphs.map(p => p.toXML());
  const docXml = XMLBuilder.createDocument(paraXmls);

  handler.addFile(DOCX_PATHS.CONTENT_TYPES, createContentTypes());
  handler.addFile(DOCX_PATHS.RELS, createRels());
  handler.addFile(DOCX_PATHS.DOCUMENT, docXml);

  await handler.save('example4-indentation.docx');
  console.log('✓ Created example4-indentation.docx');
}

/**
 * Example 5: Method chaining for fluent API
 */
async function example5MethodChaining() {
  console.log('\n=== Example 5: Method Chaining ===');

  const handler = new ZipHandler();

  // Create a fully formatted paragraph using method chaining
  const para = new Paragraph()
    .setAlignment('center')
    .setSpaceBefore(480) // 1/3 inch
    .setSpaceAfter(480)
    .addText('Fluent ')
    .addText('API ', { bold: true })
    .addText('Example', { bold: true, italic: true, color: 'FF0000' });

  const docXml = XMLBuilder.createDocument([para.toXML()]);

  handler.addFile(DOCX_PATHS.CONTENT_TYPES, createContentTypes());
  handler.addFile(DOCX_PATHS.RELS, createRels());
  handler.addFile(DOCX_PATHS.DOCUMENT, docXml);

  await handler.save('example5-method-chaining.docx');
  console.log('✓ Created example5-method-chaining.docx');
  console.log(`  Alignment: ${para.getFormatting().alignment}`);
  console.log(`  Text: "${para.getText()}"`);
}

// Helper functions to create minimal valid DOCX files
function createContentTypes(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
}

function createRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
}

// Run all examples
async function runExamples() {
  console.log('=== DocXML Paragraph & Run Examples ===');

  try {
    await example1SimpleFormattedText();
    await example2AlignmentAndSpacing();
    await example3AdvancedFormatting();
    await example4Indentation();
    await example5MethodChaining();

    console.log('\n=== All examples completed successfully! ===');
    console.log('\nGenerated files:');
    console.log('  - example1-formatted-text.docx');
    console.log('  - example2-alignment-spacing.docx');
    console.log('  - example3-advanced-formatting.docx');
    console.log('  - example4-indentation.docx');
    console.log('  - example5-method-chaining.docx');
  } catch (error) {
    console.error('Error running examples:', error);
    process.exit(1);
  }
}

// Run if executed directly
if (require.main === module) {
  runExamples();
}
