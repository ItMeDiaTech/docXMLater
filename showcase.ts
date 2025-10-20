/**
 * Complete Feature Showcase
 * Demonstrates ALL features available in docXMLater
 */

import { Document } from './src';
import { Style } from './src/formatting/Style';
import { PAGE_SIZES } from './src/utils/units';

async function createShowcaseDocument() {
  console.log('Creating comprehensive feature showcase document...\n');

  // Create a new document
  const doc = Document.create();

  // =================================================================
  // SECTION 1: STYLES
  // =================================================================
  console.log('1. Setting up custom styles...');

  // Custom title style
  const titleStyle = Style.create({
    styleId: 'ShowcaseTitle',
    name: 'Showcase Title',
    type: 'paragraph',
    basedOn: 'Normal',
    next: 'Normal',
    customStyle: true,
    paragraphFormatting: {
      alignment: 'center',
      spacing: { before: 240, after: 480 },
    },
    runFormatting: {
      font: 'Arial',
      size: 28,
      bold: true,
      color: '1F4E78',
    },
  });
  doc.addStyle(titleStyle);

  // Custom subtitle style
  const subtitleStyle = Style.create({
    styleId: 'ShowcaseSubtitle',
    name: 'Showcase Subtitle',
    type: 'paragraph',
    basedOn: 'Normal',
    next: 'Normal',
    customStyle: true,
    paragraphFormatting: {
      alignment: 'center',
      spacing: { after: 360 },
    },
    runFormatting: {
      font: 'Arial',
      size: 14,
      italic: true,
      color: '5B9BD5',
    },
  });
  doc.addStyle(subtitleStyle);

  // Body text style
  const bodyStyle = Style.create({
    styleId: 'ShowcaseBody',
    name: 'Showcase Body',
    type: 'paragraph',
    basedOn: 'Normal',
    next: 'ShowcaseBody',
    customStyle: true,
    paragraphFormatting: {
      alignment: 'justify',
      indentation: { firstLine: 720 },
      spacing: { after: 120, line: 360, lineRule: 'auto' },
    },
    runFormatting: {
      font: 'Calibri',
      size: 11,
    },
  });
  doc.addStyle(bodyStyle);

  // Code style
  const codeStyle = Style.create({
    styleId: 'ShowcaseCode',
    name: 'Showcase Code',
    type: 'paragraph',
    basedOn: 'Normal',
    next: 'ShowcaseCode',
    customStyle: true,
    paragraphFormatting: {
      spacing: { before: 80, after: 80 },
      indentation: { left: 360 },
    },
    runFormatting: {
      font: 'Courier New',
      size: 9,
      color: '2D572C',
    },
  });
  doc.addStyle(codeStyle);

  // Quote style
  const quoteStyle = Style.create({
    styleId: 'ShowcaseQuote',
    name: 'Showcase Quote',
    type: 'paragraph',
    basedOn: 'Normal',
    next: 'Normal',
    customStyle: true,
    paragraphFormatting: {
      alignment: 'left',
      indentation: { left: 720, right: 720 },
      spacing: { before: 120, after: 120 },
    },
    runFormatting: {
      font: 'Georgia',
      size: 11,
      italic: true,
      color: '666666',
    },
  });
  doc.addStyle(quoteStyle);

  // Add standard heading styles
  doc.addStyle(Style.createHeadingStyle(1));
  doc.addStyle(Style.createHeadingStyle(2));
  doc.addStyle(Style.createHeadingStyle(3));

  // =================================================================
  // SECTION 2: PAGE SETUP
  // =================================================================
  console.log('2. Configuring page setup...');

  doc.getSection()
    .setPageSize(PAGE_SIZES.LETTER.width, PAGE_SIZES.LETTER.height, 'portrait')
    .setMargins({
      top: 1440,    // 1 inch
      bottom: 1440,
      left: 1440,
      right: 1440,
      header: 720,
      footer: 720,
    })
    .setPageNumbering(1, 'decimal');

  // =================================================================
  // SECTION 3: DOCUMENT HEADER
  // =================================================================
  console.log('3. Creating document header...');

  const title = doc.createParagraph('docXMLater Complete Feature Showcase');
  title.setStyle('ShowcaseTitle');

  const subtitle = doc.createParagraph('A Comprehensive Demonstration of All Capabilities');
  subtitle.setStyle('ShowcaseSubtitle');

  doc.createParagraph(); // Spacing

  const intro = doc.createParagraph(
    'This document demonstrates every feature available in the docXMLater library, ' +
    'including text formatting, paragraph styling, custom styles, tables, lists, ' +
    'sections, and more. Each section below showcases a specific capability with ' +
    'working examples.'
  );
  intro.setStyle('ShowcaseBody');

  // =================================================================
  // SECTION 4: TEXT FORMATTING
  // =================================================================
  console.log('4. Demonstrating text formatting...');

  doc.createParagraph(); // Spacing
  const textHeading = doc.createParagraph('1. Text Formatting');
  textHeading.setStyle('Heading1');

  const para1 = doc.createParagraph();
  para1.addText('This paragraph demonstrates various text formatting options: ');
  para1.addText('bold text', { bold: true });
  para1.addText(', ');
  para1.addText('italic text', { italic: true });
  para1.addText(', ');
  para1.addText('underlined text', { underline: 'single' });
  para1.addText(', ');
  para1.addText('strikethrough text', { strike: true });
  para1.addText(', ');
  para1.addText('subscript', { subscript: true });
  para1.addText(' and ');
  para1.addText('superscript', { superscript: true });
  para1.addText('.');

  const para2 = doc.createParagraph();
  para2.addText('Font variations: ');
  para2.addText('Arial font', { font: 'Arial' });
  para2.addText(', ');
  para2.addText('Times New Roman', { font: 'Times New Roman' });
  para2.addText(', ');
  para2.addText('Courier New', { font: 'Courier New' });
  para2.addText('.');

  const para3 = doc.createParagraph();
  para3.addText('Size variations: ');
  para3.addText('8pt', { size: 8 });
  para3.addText(', ');
  para3.addText('12pt', { size: 12 });
  para3.addText(', ');
  para3.addText('18pt', { size: 18 });
  para3.addText(', ');
  para3.addText('24pt', { size: 24 });
  para3.addText('.');

  const para4 = doc.createParagraph();
  para4.addText('Color examples: ');
  para4.addText('red', { color: 'FF0000' });
  para4.addText(', ');
  para4.addText('green', { color: '00FF00' });
  para4.addText(', ');
  para4.addText('blue', { color: '0000FF' });
  para4.addText(', ');
  para4.addText('purple', { color: '800080' });
  para4.addText('.');

  const para5 = doc.createParagraph();
  para5.addText('Highlighting: ');
  para5.addText('yellow highlight', { highlight: 'yellow' });
  para5.addText(', ');
  para5.addText('cyan highlight', { highlight: 'cyan' });
  para5.addText(', ');
  para5.addText('lightGray highlight', { highlight: 'lightGray' });
  para5.addText('.');

  const para6 = doc.createParagraph();
  para6.addText('Text effects: ');
  para6.addText('SMALL CAPS', { smallCaps: true });
  para6.addText(', ');
  para6.addText('ALL CAPS', { allCaps: true });
  para6.addText('.');

  const para7 = doc.createParagraph();
  para7.addText('Combinations: ');
  para7.addText('Bold + Italic + Underline + Red', {
    bold: true,
    italic: true,
    underline: 'single',
    color: 'FF0000',
  });
  para7.addText(', ');
  para7.addText('Large + Bold + Blue + Yellow Highlight', {
    size: 14,
    bold: true,
    color: '0000FF',
    highlight: 'yellow',
  });
  para7.addText('.');

  // =================================================================
  // SECTION 5: PARAGRAPH FORMATTING
  // =================================================================
  console.log('5. Demonstrating paragraph formatting...');

  doc.createParagraph(); // Spacing
  const paraHeading = doc.createParagraph('2. Paragraph Formatting');
  paraHeading.setStyle('Heading1');

  const alignHeading = doc.createParagraph('Alignment Options:');
  alignHeading.setStyle('Heading2');

  const leftAlign = doc.createParagraph('This paragraph is left-aligned (default).');
  leftAlign.setAlignment('left');

  const centerAlign = doc.createParagraph('This paragraph is center-aligned.');
  centerAlign.setAlignment('center');

  const rightAlign = doc.createParagraph('This paragraph is right-aligned.');
  rightAlign.setAlignment('right');

  const justifyAlign = doc.createParagraph(
    'This paragraph is justified. When text wraps to multiple lines, it spreads ' +
    'evenly across the line width, creating clean left and right edges. This is ' +
    'commonly used in professional documents and books.'
  );
  justifyAlign.setAlignment('justify');

  doc.createParagraph(); // Spacing
  const indentHeading = doc.createParagraph('Indentation:');
  indentHeading.setStyle('Heading2');

  const firstLineIndent = doc.createParagraph(
    'This paragraph has a first-line indent of 0.5 inches, commonly used for body ' +
    'text in formal documents and books. The subsequent lines remain at the left margin.'
  );
  firstLineIndent.setFirstLineIndent(720);

  const hangingIndent = doc.createParagraph(
    'This paragraph has a hanging indent where the first line extends to the left ' +
    'margin and subsequent lines are indented. This is commonly used for bibliographies ' +
    'and reference lists.'
  );
  hangingIndent.setLeftIndent(720);
  // Note: Hanging indent requires using the formatting property directly
  if (!hangingIndent['formatting'].indentation) {
    hangingIndent['formatting'].indentation = {};
  }
  hangingIndent['formatting'].indentation.hanging = 720;
  // Clear firstLine if set to avoid conflicts
  delete hangingIndent['formatting'].indentation.firstLine;

  const leftIndent = doc.createParagraph(
    'This entire paragraph is indented from the left margin by 1 inch, creating a ' +
    'block quote effect.'
  );
  leftIndent.setLeftIndent(1440);

  const bothIndent = doc.createParagraph(
    'This paragraph is indented from both left and right margins, creating a centered ' +
    'block effect often used for special callouts or quotations.'
  );
  bothIndent.setLeftIndent(1440);
  bothIndent.setRightIndent(1440);

  doc.createParagraph(); // Spacing
  const spacingHeading = doc.createParagraph('Spacing:');
  spacingHeading.setStyle('Heading2');

  const spaceBefore = doc.createParagraph('This paragraph has 240 twips (1/6 inch) spacing before it.');
  spaceBefore.setSpaceBefore(240);

  const spaceAfter = doc.createParagraph('This paragraph has 240 twips (1/6 inch) spacing after it.');
  spaceAfter.setSpaceAfter(240);

  const lineSpacing = doc.createParagraph(
    'This paragraph has double line spacing. Line spacing controls the vertical ' +
    'distance between lines of text within a paragraph. Double spacing is often ' +
    'required for academic papers and draft documents.'
  );
  lineSpacing.setLineSpacing(480, 'auto');

  // =================================================================
  // SECTION 6: NUMBERED LISTS
  // =================================================================
  console.log('6. Creating numbered lists...');

  doc.createParagraph(); // Spacing
  const numberedHeading = doc.createParagraph('3. Numbered Lists');
  numberedHeading.setStyle('Heading1');

  const numberedIntro = doc.createParagraph(
    'The library supports multi-level numbered lists with automatic numbering:'
  );
  numberedIntro.setStyle('ShowcaseBody');

  const numberedList = doc.createNumberedList();

  const num1 = doc.createParagraph('First main item in the numbered list');
  num1.setNumbering(numberedList, 0);

  const num1a = doc.createParagraph('First sub-item under item 1');
  num1a.setNumbering(numberedList, 1);

  const num1b = doc.createParagraph('Second sub-item under item 1');
  num1b.setNumbering(numberedList, 1);

  const num1b1 = doc.createParagraph('Third-level item under 1.b');
  num1b1.setNumbering(numberedList, 2);

  const num1b2 = doc.createParagraph('Another third-level item under 1.b');
  num1b2.setNumbering(numberedList, 2);

  const num1c = doc.createParagraph('Third sub-item under item 1');
  num1c.setNumbering(numberedList, 1);

  const num2 = doc.createParagraph('Second main item in the numbered list');
  num2.setNumbering(numberedList, 0);

  const num2a = doc.createParagraph('First sub-item under item 2');
  num2a.setNumbering(numberedList, 1);

  const num2b = doc.createParagraph('Second sub-item under item 2');
  num2b.setNumbering(numberedList, 1);

  const num3 = doc.createParagraph('Third main item in the numbered list');
  num3.setNumbering(numberedList, 0);

  // =================================================================
  // SECTION 7: BULLETED LISTS
  // =================================================================
  console.log('7. Creating bulleted lists...');

  doc.createParagraph(); // Spacing
  const bulletHeading = doc.createParagraph('4. Bulleted Lists');
  bulletHeading.setStyle('Heading1');

  const bulletIntro = doc.createParagraph(
    'Bulleted lists are perfect for unordered information and feature hierarchies:'
  );
  bulletIntro.setStyle('ShowcaseBody');

  const bulletList = doc.createBulletList();

  const bullet1 = doc.createParagraph('Primary feature category');
  bullet1.setNumbering(bulletList, 0);

  const bullet1a = doc.createParagraph('Supporting detail for primary feature');
  bullet1a.setNumbering(bulletList, 1);

  const bullet1b = doc.createParagraph('Additional supporting detail');
  bullet1b.setNumbering(bulletList, 1);

  const bullet1b1 = doc.createParagraph('Fine-grained detail');
  bullet1b1.setNumbering(bulletList, 2);

  const bullet2 = doc.createParagraph('Another primary feature category');
  bullet2.setNumbering(bulletList, 0);

  const bullet2a = doc.createParagraph('Supporting detail for second feature');
  bullet2a.setNumbering(bulletList, 1);

  const bullet3 = doc.createParagraph('Third primary feature category');
  bullet3.setNumbering(bulletList, 0);

  // =================================================================
  // SECTION 8: TABLES
  // =================================================================
  console.log('8. Creating tables...');

  doc.createParagraph(); // Spacing
  const tableHeading = doc.createParagraph('5. Tables');
  tableHeading.setStyle('Heading1');

  const tableIntro = doc.createParagraph(
    'Tables support formatting, borders, shading, cell merging, and alignment:'
  );
  tableIntro.setStyle('ShowcaseBody');

  // Simple table
  const simpleTableHeading = doc.createParagraph('Basic Table:');
  simpleTableHeading.setStyle('Heading2');

  const simpleTable = doc.createTable(4, 3);
  simpleTable
    .setWidth(7200)
    .setAlignment('center')
    .setAllBorders({ style: 'single', size: 4, color: '000000' });

  // Header row
  simpleTable.getRow(0)?.getCell(0)?.createParagraph('Product');
  simpleTable.getRow(0)?.getCell(1)?.createParagraph('Price');
  simpleTable.getRow(0)?.getCell(2)?.createParagraph('Quantity');
  simpleTable.getRow(0)?.setHeader(true);

  for (let i = 0; i < 3; i++) {
    simpleTable.getRow(0)?.getCell(i)?.setShading({ fill: 'D9E2F3' });
    simpleTable.getRow(0)?.getCell(i)?.getParagraphs()[0]?.getRuns()[0]?.setBold(true);
    simpleTable.getRow(0)?.getCell(i)?.getParagraphs()[0]?.setAlignment('center');
  }

  // Data rows
  simpleTable.getRow(1)?.getCell(0)?.createParagraph('Widget A');
  simpleTable.getRow(1)?.getCell(1)?.createParagraph('$29.99');
  simpleTable.getRow(1)?.getCell(2)?.createParagraph('150');

  simpleTable.getRow(2)?.getCell(0)?.createParagraph('Widget B');
  simpleTable.getRow(2)?.getCell(1)?.createParagraph('$39.99');
  simpleTable.getRow(2)?.getCell(2)?.createParagraph('200');

  simpleTable.getRow(3)?.getCell(0)?.createParagraph('Widget C');
  simpleTable.getRow(3)?.getCell(1)?.createParagraph('$49.99');
  simpleTable.getRow(3)?.getCell(2)?.createParagraph('175');

  // Complex table with merged cells
  doc.createParagraph(); // Spacing
  const complexTableHeading = doc.createParagraph('Advanced Table with Merged Cells:');
  complexTableHeading.setStyle('Heading2');

  const complexTable = doc.createTable(6, 4);
  complexTable
    .setWidth(8000)
    .setAlignment('center')
    .setAllBorders({ style: 'single', size: 6, color: '2F5496' });

  // Title row with merged cells
  const titleRow = complexTable.getRow(0);
  titleRow?.getCell(0)?.createParagraph('2024 Sales Performance Report');
  titleRow?.getCell(0)?.setColumnSpan(4);
  titleRow?.getCell(0)?.setShading({ fill: '2F5496' });
  titleRow?.getCell(0)?.getParagraphs()[0]?.setAlignment('center');
  titleRow?.getCell(0)?.getParagraphs()[0]?.getRuns()[0]?.setBold(true).setColor('FFFFFF').setSize(14);
  for (let i = 1; i < 4; i++) {
    titleRow?.getCell(i)?.setWidth(0);
  }

  // Header row
  const headerRow = complexTable.getRow(1);
  headerRow?.getCell(0)?.createParagraph('Region');
  headerRow?.getCell(1)?.createParagraph('Q1');
  headerRow?.getCell(2)?.createParagraph('Q2');
  headerRow?.getCell(3)?.createParagraph('Total');
  headerRow?.setHeader(true);

  for (let i = 0; i < 4; i++) {
    headerRow?.getCell(i)?.setShading({ fill: 'D9E2F3' });
    headerRow?.getCell(i)?.getParagraphs()[0]?.getRuns()[0]?.setBold(true);
    headerRow?.getCell(i)?.getParagraphs()[0]?.setAlignment('center');
  }

  // Data rows with alternating colors
  const regions = ['North', 'South', 'East', 'West'];
  const q1Sales = ['$2.5M', '$3.1M', '$2.8M', '$3.3M'];
  const q2Sales = ['$2.9M', '$3.4M', '$3.1M', '$3.6M'];
  const totals = ['$5.4M', '$6.5M', '$5.9M', '$6.9M'];

  for (let i = 0; i < 4; i++) {
    const row = complexTable.getRow(i + 2);
    const bgColor = i % 2 === 0 ? 'F2F2F2' : undefined;

    row?.getCell(0)?.createParagraph(regions[i]!);
    row?.getCell(1)?.createParagraph(q1Sales[i]!);
    row?.getCell(2)?.createParagraph(q2Sales[i]!);
    row?.getCell(3)?.createParagraph(totals[i]!);

    if (bgColor) {
      for (let j = 0; j < 4; j++) {
        row?.getCell(j)?.setShading({ fill: bgColor });
      }
    }

    // Center align numbers
    for (let j = 1; j < 4; j++) {
      row?.getCell(j)?.getParagraphs()[0]?.setAlignment('center');
    }
  }

  // =================================================================
  // SECTION 9: CUSTOM STYLES IN ACTION
  // =================================================================
  console.log('9. Demonstrating custom styles...');

  doc.createParagraph(); // Spacing
  const styleHeading = doc.createParagraph('6. Custom Styles');
  styleHeading.setStyle('Heading1');

  const styleIntro = doc.createParagraph(
    'Custom styles ensure consistent formatting throughout your document. Here are examples:'
  );
  styleIntro.setStyle('ShowcaseBody');

  const codeHeading = doc.createParagraph('Code Example:');
  codeHeading.setStyle('Heading2');

  const code1 = doc.createParagraph('const doc = Document.create();');
  code1.setStyle('ShowcaseCode');

  const code2 = doc.createParagraph('const para = doc.createParagraph("Hello World");');
  code2.setStyle('ShowcaseCode');

  const code3 = doc.createParagraph('para.addText("Bold text", { bold: true });');
  code3.setStyle('ShowcaseCode');

  const code4 = doc.createParagraph('await doc.save("output.docx");');
  code4.setStyle('ShowcaseCode');

  doc.createParagraph(); // Spacing
  const quoteHeading = doc.createParagraph('Quote Example:');
  quoteHeading.setStyle('Heading2');

  const quote = doc.createParagraph(
    'The best code is the code you don\'t write. Simplicity is the ultimate sophistication. ' +
    'Focus on solving real problems rather than imaginary ones.'
  );
  quote.setStyle('ShowcaseQuote');

  // =================================================================
  // SECTION 10: ADVANCED FEATURES
  // =================================================================
  console.log('10. Showcasing advanced features...');

  doc.createParagraph(); // Spacing
  const advancedHeading = doc.createParagraph('7. Advanced Features');
  advancedHeading.setStyle('Heading1');

  const featuresIntro = doc.createParagraph(
    'The library includes sophisticated features for professional document creation:'
  );
  featuresIntro.setStyle('ShowcaseBody');

  const featuresList = doc.createBulletList();

  const feat1 = doc.createParagraph('ZIP archive handling with 14 helper methods');
  feat1.setNumbering(featuresList, 0);

  const feat2 = doc.createParagraph('Full UTF-8 encoding support for international characters');
  feat2.setNumbering(featuresList, 0);

  const feat3 = doc.createParagraph('XML generation compliant with ECMA-376 standards');
  feat3.setNumbering(featuresList, 0);

  const feat4 = doc.createParagraph('Style inheritance and cascading');
  feat4.setNumbering(featuresList, 0);

  const feat5 = doc.createParagraph('Multi-level numbering with 9 levels supported');
  feat5.setNumbering(featuresList, 0);

  const feat6 = doc.createParagraph('Table cell spanning and merging');
  feat6.setNumbering(featuresList, 0);

  const feat7 = doc.createParagraph('Section configuration (page size, margins, orientation)');
  feat7.setNumbering(featuresList, 0);

  const feat8 = doc.createParagraph('Comprehensive error handling and validation');
  feat8.setNumbering(featuresList, 0);

  // =================================================================
  // SECTION 11: STATISTICS AND INFO
  // =================================================================
  console.log('11. Adding document statistics...');

  doc.createParagraph(); // Spacing
  const statsHeading = doc.createParagraph('8. Document Statistics');
  statsHeading.setStyle('Heading1');

  const statsTable = doc.createTable(6, 2);
  statsTable
    .setWidth(6000)
    .setAlignment('center')
    .setAllBorders({ style: 'single', size: 4, color: '4472C4' });

  // Header
  statsTable.getRow(0)?.getCell(0)?.createParagraph('Metric');
  statsTable.getRow(0)?.getCell(0)?.setShading({ fill: '4472C4' });
  statsTable.getRow(0)?.getCell(0)?.getParagraphs()[0]?.getRuns()[0]?.setBold(true).setColor('FFFFFF');
  statsTable.getRow(0)?.getCell(0)?.getParagraphs()[0]?.setAlignment('center');

  statsTable.getRow(0)?.getCell(1)?.createParagraph('Value');
  statsTable.getRow(0)?.getCell(1)?.setShading({ fill: '4472C4' });
  statsTable.getRow(0)?.getCell(1)?.getParagraphs()[0]?.getRuns()[0]?.setBold(true).setColor('FFFFFF');
  statsTable.getRow(0)?.getCell(1)?.getParagraphs()[0]?.setAlignment('center');

  // Data
  const statsData = [
    ['Library Version', 'v0.27.0'],
    ['Total Test Suite', '226+ tests'],
    ['Source Files', '48 files'],
    ['Lines of Code', '~10,000+'],
    ['Test Coverage', '>90%'],
  ];

  for (let i = 0; i < statsData.length; i++) {
    const row = statsTable.getRow(i + 1);
    row?.getCell(0)?.createParagraph(statsData[i]![0]!);
    row?.getCell(1)?.createParagraph(statsData[i]![1]!);

    if (i % 2 === 0) {
      row?.getCell(0)?.setShading({ fill: 'F2F2F2' });
      row?.getCell(1)?.setShading({ fill: 'F2F2F2' });
    }

    row?.getCell(1)?.getParagraphs()[0]?.setAlignment('center');
  }

  // =================================================================
  // SECTION 12: CONCLUSION
  // =================================================================
  console.log('12. Adding conclusion...');

  doc.createParagraph(); // Spacing
  const conclusionHeading = doc.createParagraph('Conclusion');
  conclusionHeading.setStyle('Heading1');

  const conclusion = doc.createParagraph(
    'The docXMLater library provides a comprehensive, production-ready solution for ' +
    'programmatic DOCX document creation and manipulation. With support for advanced ' +
    'formatting, custom styles, complex tables, multi-level lists, and section ' +
    'configuration, it enables developers to create professional Word documents entirely ' +
    'through code. The library follows ECMA-376 standards and maintains compatibility ' +
    'with Microsoft Word 2016 and later versions.'
  );
  conclusion.setStyle('ShowcaseBody');

  doc.createParagraph(); // Spacing
  const finalNote = doc.createParagraph(
    'This document itself was generated programmatically using docXMLater, demonstrating ' +
    'the library\'s capabilities in a real-world example. All formatting, styles, tables, ' +
    'and lists were created through the API without any manual intervention.'
  );
  finalNote.setStyle('ShowcaseBody');

  // =================================================================
  // SAVE DOCUMENT
  // =================================================================
  console.log('\n13. Saving document...');
  await doc.save('showcase.docx');
  console.log('Document saved as showcase.docx');

  // Display final statistics
  console.log('\nFinal Document Statistics:');
  console.log(`  Paragraphs: ${doc.getParagraphCount()}`);
  console.log(`  Tables: ${doc.getTableCount()}`);
  console.log(`  Custom Styles: ${doc.getStyles().length}`);
  console.log(`  Word Count: ${doc.getWordCount()}`);
  console.log('\nShowcase document created successfully!');
}

// Run the showcase
createShowcaseDocument().catch((error) => {
  console.error('Error creating showcase document:', error);
  process.exit(1);
});
