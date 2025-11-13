/**
 * Bookmark Usage Examples
 *
 * Demonstrates how to use bookmarks for internal navigation in documents.
 * Bookmarks mark specific locations and can be referenced by internal hyperlinks.
 */

import { Document, Hyperlink } from '../../src';
import * as path from 'path';
import * as fs from 'fs';

// Create output directory if it doesn't exist
const outputDir = __dirname;
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

/**
 * Example 1: Simple bookmark with internal hyperlink
 */
async function example1_SimpleBookmark() {
  console.log('Example 1: Simple bookmark with navigation...');

  const doc = Document.create({
    properties: {
      title: 'Simple Bookmark Example',
      creator: 'DocXML Examples',
    },
  });

  // Create a bookmark at a specific paragraph
  const bookmark = doc.createBookmark('important_section');

  // Add a paragraph with a link to the bookmark
  doc.createParagraph()
    .addText('Click ')
    .addHyperlink(
      Hyperlink.createInternal('important_section', 'here', { color: '0000FF', underline: 'single' })
    )
    .addText(' to jump to the important section.');

  // Add some filler content
  for (let i = 0; i < 10; i++) {
    doc.createParagraph(`This is filler paragraph ${i + 1}.`)
      .setSpaceAfter(120);
  }

  // Add the bookmarked paragraph
  const targetParagraph = doc.createParagraph('Important Section: This paragraph is bookmarked!')
    .setStyle('Heading1');

  // Add the bookmark to the paragraph
  targetParagraph.addBookmark(bookmark);

  // Add more content after
  doc.createParagraph(
    'This is the important content that you jumped to. You can use bookmarks ' +
    'to create internal navigation within your documents, making it easy for ' +
    'readers to jump to specific sections.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Save document
  const outputPath = path.join(outputDir, 'example1-simple-bookmark.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log('  Note: Click the link in the document to jump to the bookmarked section!');
}

/**
 * Example 2: Multiple bookmarks with navigation menu
 */
async function example2_NavigationMenu() {
  console.log('Example 2: Multiple bookmarks with navigation menu...');

  const doc = Document.create({
    properties: {
      title: 'Navigation Menu Example',
      creator: 'DocXML Examples',
    },
  });

  // Create bookmarks for different sections
  const introBookmark = doc.createBookmark('introduction');
  const methodsBookmark = doc.createBookmark('methods');
  const resultsBookmark = doc.createBookmark('results');
  const conclusionBookmark = doc.createBookmark('conclusion');

  // Create navigation menu
  doc.createParagraph('Document Navigation')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(240);

  doc.createParagraph()
    .addText('• ')
    .addHyperlink(Hyperlink.createInternal('introduction', 'Introduction'))
    .setSpaceBefore(120);

  doc.createParagraph()
    .addText('• ')
    .addHyperlink(Hyperlink.createInternal('methods', 'Methods'))
    .setSpaceBefore(120);

  doc.createParagraph()
    .addText('• ')
    .addHyperlink(Hyperlink.createInternal('results', 'Results'))
    .setSpaceBefore(120);

  doc.createParagraph()
    .addText('• ')
    .addHyperlink(Hyperlink.createInternal('conclusion', 'Conclusion'))
    .setSpaceBefore(120)
    .setSpaceAfter(480);

  // Add some page breaks to make navigation more dramatic
  doc.createParagraph('').setPageBreakBefore(true);

  // Add sections with bookmarks
  const introPara = doc.createParagraph('1. Introduction')
    .setStyle('Heading1');
  introPara.addBookmark(introBookmark);

  doc.createParagraph(
    'This is the introduction section. '.repeat(20)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('').setPageBreakBefore(true);

  const methodsPara = doc.createParagraph('2. Methods')
    .setStyle('Heading1');
  methodsPara.addBookmark(methodsBookmark);

  doc.createParagraph(
    'This is the methods section with detailed methodology. '.repeat(20)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('').setPageBreakBefore(true);

  const resultsPara = doc.createParagraph('3. Results')
    .setStyle('Heading1');
  resultsPara.addBookmark(resultsBookmark);

  doc.createParagraph(
    'This is the results section presenting our findings. '.repeat(20)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('').setPageBreakBefore(true);

  const conclusionPara = doc.createParagraph('4. Conclusion')
    .setStyle('Heading1');
  conclusionPara.addBookmark(conclusionBookmark);

  doc.createParagraph(
    'This is the conclusion section summarizing everything. '.repeat(20)
  )
    .setAlignment('justify');

  // Save document
  const outputPath = path.join(outputDir, 'example2-navigation-menu.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 3: Heading bookmarks for automatic TOC-style links
 */
async function example3_HeadingBookmarks() {
  console.log('Example 3: Automatic heading bookmarks...');

  const doc = Document.create({
    properties: {
      title: 'Heading Bookmarks Example',
      creator: 'DocXML Examples',
    },
  });

  // Create title
  doc.createParagraph('Research Paper')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  // Create "manual TOC" using heading bookmarks
  doc.createParagraph('Contents')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  const headings = [
    'Abstract',
    'Background',
    'Literature Review',
    'Methodology',
    'Data Analysis',
    'Findings',
    'Discussion',
    'Conclusion',
    'References',
  ];

  // Create bookmarks for all headings
  const bookmarks = headings.map(heading =>
    doc.createHeadingBookmark(heading)
  );

  // Create clickable links in the "TOC"
  headings.forEach((heading, index) => {
    const bookmark = bookmarks[index];
    if (!bookmark) return;

    doc.createParagraph()
      .addText(`${index + 1}. `)
      .addHyperlink(
        Hyperlink.createInternal(
          bookmark.getName(),
          heading,
          { color: '1F4D78', underline: 'single' }
        )
      )
      .setSpaceBefore(120)
      .setLeftIndent(360);
  });

  doc.createParagraph('').setSpaceAfter(480);

  // Add actual sections with bookmarks
  headings.forEach((heading, index) => {
    const bookmark = bookmarks[index];
    if (!bookmark) return;

    if (index > 0) {
      doc.createParagraph('').setPageBreakBefore(true);
    }

    const headingPara = doc.createParagraph(`${index + 1}. ${heading}`)
      .setStyle('Heading1');
    headingPara.addBookmark(bookmark);

    doc.createParagraph(
      `This is the ${heading.toLowerCase()} section with detailed content. `.repeat(15)
    )
      .setAlignment('justify')
      .setSpaceAfter(240);
  });

  // Save document
  const outputPath = path.join(outputDir, 'example3-heading-bookmarks.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 4: Cross-references using bookmarks
 */
async function example4_CrossReferences() {
  console.log('Example 4: Cross-references with bookmarks...');

  const doc = Document.create({
    properties: {
      title: 'Cross-References Example',
      creator: 'DocXML Examples',
    },
  });

  // Create bookmarks for figures and tables
  const figure1 = doc.createBookmark('figure_1');
  const table1 = doc.createBookmark('table_1');
  const appendixA = doc.createBookmark('appendix_a');

  doc.createParagraph('Technical Documentation')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  // Main content with cross-references
  doc.createParagraph('Introduction')
    .setStyle('Heading1');

  doc.createParagraph()
    .addText('This document contains several important elements. ')
    .addText('See ')
    .addHyperlink(Hyperlink.createInternal('figure_1', 'Figure 1', { bold: true }))
    .addText(' for the system architecture diagram. ')
    .addText('The performance metrics are detailed in ')
    .addHyperlink(Hyperlink.createInternal('table_1', 'Table 1', { bold: true }))
    .addText('. Additional information can be found in ')
    .addHyperlink(Hyperlink.createInternal('appendix_a', 'Appendix A', { bold: true }))
    .addText('.')
    .setAlignment('justify')
    .setSpaceAfter(480);

  // Add more content
  doc.createParagraph('System Overview')
    .setStyle('Heading1');

  doc.createParagraph(
    'The system architecture follows a modular design pattern. '.repeat(10)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Figure 1 with bookmark
  doc.createParagraph('').setPageBreakBefore(true);
  const fig1Para = doc.createParagraph('Figure 1: System Architecture Diagram')
    .setAlignment('center')
    .setStyle('Heading2');
  fig1Para.addBookmark(figure1);

  doc.createParagraph('[System architecture diagram would go here]')
    .setAlignment('center')
    .setSpaceAfter(480);

  // Table 1 with bookmark
  doc.createParagraph('').setPageBreakBefore(true);
  const table1Para = doc.createParagraph('Table 1: Performance Metrics')
    .setAlignment('center')
    .setStyle('Heading2');
  table1Para.addBookmark(table1);

  const table = doc.createTable(4, 3);

  // Header row
  table.getRow(0)?.getCell(0)?.createParagraph().addText('Metric', { bold: true });
  table.getRow(0)?.getCell(1)?.createParagraph().addText('Value', { bold: true });
  table.getRow(0)?.getCell(2)?.createParagraph().addText('Unit', { bold: true });
  table.getRow(0)?.setHeader(true);

  // Data rows
  table.getRow(1)?.getCell(0)?.createParagraph().addText('Response Time');
  table.getRow(1)?.getCell(1)?.createParagraph().addText('45');
  table.getRow(1)?.getCell(2)?.createParagraph().addText('ms');

  table.getRow(2)?.getCell(0)?.createParagraph().addText('Throughput');
  table.getRow(2)?.getCell(1)?.createParagraph().addText('1000');
  table.getRow(2)?.getCell(2)?.createParagraph().addText('req/s');

  table.getRow(3)?.getCell(0)?.createParagraph().addText('Error Rate');
  table.getRow(3)?.getCell(1)?.createParagraph().addText('0.01');
  table.getRow(3)?.getCell(2)?.createParagraph().addText('%');

  doc.createParagraph('').setSpaceAfter(480);

  // Appendix with bookmark
  doc.createParagraph('').setPageBreakBefore(true);
  const appendixPara = doc.createParagraph('Appendix A: Additional Details')
    .setStyle('Heading1');
  appendixPara.addBookmark(appendixA);

  doc.createParagraph(
    'This appendix contains supplementary information referenced earlier in the document. '.repeat(15)
  )
    .setAlignment('justify');

  // Save document
  const outputPath = path.join(outputDir, 'example4-cross-references.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 5: Bookmark name normalization and management
 */
async function example5_BookmarkManagement() {
  console.log('Example 5: Bookmark management and naming...');

  const doc = Document.create({
    properties: {
      title: 'Bookmark Management Example',
      creator: 'DocXML Examples',
    },
  });

  doc.createParagraph('Bookmark Name Normalization Demo')
    .setStyle('Title')
    .setAlignment('center')
    .setSpaceAfter(480);

  doc.createParagraph(
    'This example demonstrates how bookmark names are automatically normalized ' +
    'to follow Word\'s naming rules. Bookmark names must start with a letter or ' +
    'underscore, can only contain letters, numbers, and underscores, and are ' +
    'limited to 40 characters.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Create bookmarks with various names (will be normalized automatically)
  const testNames = [
    'Valid_Bookmark_123',           // Valid - no change
    'Invalid Name With Spaces',     // Will become: Invalid_Name_With_Spaces
    '123_starts_with_number',       // Will become: _123_starts_with_number
    'special@#$characters!',        // Will become: special___characters_
    'very-long-bookmark-name-that-exceeds-forty-characters-limit',  // Will be truncated
  ];

  doc.createParagraph('Test Bookmark Names:')
    .setStyle('Heading2')
    .setSpaceBefore(240)
    .setSpaceAfter(120);

  const bookmarks = testNames.map(name => doc.createBookmark(name));

  // Show original and normalized names
  testNames.forEach((originalName, index) => {
    const bookmark = bookmarks[index];
    if (!bookmark) return;

    doc.createParagraph()
      .addText(`Original: "${originalName}"`, { font: 'Courier New', size: 10 })
      .setSpaceBefore(120);

    doc.createParagraph()
      .addText(`Normalized: "${bookmark.getName()}"`, { font: 'Courier New', size: 10, color: '008000' })
      .setSpaceBefore(60)
      .setSpaceAfter(120);

    // Add bookmarked content
    const para = doc.createParagraph(`Section: ${bookmark.getName()}`)
      .setStyle('Heading3')
      .setSpaceBefore(240);
    para.addBookmark(bookmark);

    doc.createParagraph('This section is bookmarked.')
      .setSpaceAfter(240);
  });

  // Demonstrate duplicate handling
  doc.createParagraph('Duplicate Name Handling:')
    .setStyle('Heading2')
    .setSpaceBefore(480)
    .setSpaceAfter(120);

  doc.createParagraph(
    'The bookmark manager automatically ensures all bookmark names are unique. ' +
    'If you try to create a bookmark with a name that already exists, it will ' +
    'add a suffix to make it unique.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Try to create duplicate bookmarks
  const dup1 = doc.createBookmark('section_1');
  const dup2 = doc.createBookmark('section_1'); // Will become section_1_1
  const dup3 = doc.createBookmark('section_1'); // Will become section_1_2

  doc.createParagraph()
    .addText(`First bookmark: "${dup1.getName()}"`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(120);
  doc.createParagraph()
    .addText(`Second bookmark: "${dup2.getName()}"`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(60);
  doc.createParagraph()
    .addText(`Third bookmark: "${dup3.getName()}"`, { font: 'Courier New', size: 10 })
    .setSpaceBefore(60);

  // Show bookmark stats
  doc.createParagraph('').setPageBreakBefore(true);
  doc.createParagraph('Bookmark Statistics')
    .setStyle('Heading2')
    .setSpaceAfter(120);

  const stats = doc.getBookmarkManager().getStats();
  doc.createParagraph()
    .addText(`Total bookmarks: ${stats.total}`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(120);
  doc.createParagraph()
    .addText(`Next ID: ${stats.nextId}`, { font: 'Courier New', size: 11 })
    .setSpaceBefore(60);

  doc.createParagraph()
    .addText('All bookmark names:', { font: 'Courier New', size: 11 })
    .setSpaceBefore(120)
    .setSpaceAfter(60);

  stats.names.forEach((name, index) => {
    doc.createParagraph()
      .addText(`  ${index + 1}. ${name}`, { font: 'Courier New', size: 10, color: '666666' })
      .setSpaceBefore(40);
  });

  // Save document
  const outputPath = path.join(outputDir, 'example5-bookmark-management.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Main function to run all examples
 */
async function main() {
  console.log('Running Bookmark Examples...\n');

  try {
    await example1_SimpleBookmark();
    await example2_NavigationMenu();
    await example3_HeadingBookmarks();
    await example4_CrossReferences();
    await example5_BookmarkManagement();

    console.log('\n✓ All examples completed successfully!');
    console.log(`\nOutput files saved to: ${outputDir}`);
    console.log('\n📝 Important: Open the documents in Microsoft Word and click the');
    console.log('   hyperlinks to test bookmark navigation!');
  } catch (error) {
    console.error('Error running examples:', error);
    process.exit(1);
  }
}

// Run examples if executed directly
if (require.main === module) {
  main();
}

export {
  example1_SimpleBookmark,
  example2_NavigationMenu,
  example3_HeadingBookmarks,
  example4_CrossReferences,
  example5_BookmarkManagement,
};
