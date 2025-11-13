/**
 * Header and Footer Usage Examples
 *
 * This example demonstrates how to add headers, footers, and dynamic fields to documents.
 * Includes page numbers, dates, and different first page layouts.
 */

import { Document, Header, Footer, Field } from '../../src';
import * as path from 'path';
import * as fs from 'fs';

// Create output directory if it doesn't exist
const outputDir = __dirname;
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

/**
 * Example 1: Simple header and footer with page numbers
 */
async function example1_SimpleHeaderFooter() {
  console.log('Example 1: Simple header and footer with page numbers...');

  const doc = Document.create({
    properties: {
      title: 'Simple Header/Footer Example',
      creator: 'DocXML Examples',
    },
  });

  // Create header with title
  const header = Header.createDefault();
  const headerPara = header.createParagraph();
  headerPara.setAlignment('right');
  headerPara.addText('Document Title', { bold: true });

  // Create footer with page numbers
  const footer = Footer.createDefault();
  const footerPara = footer.createParagraph();
  footerPara.setAlignment('center');
  footerPara.addText('Page ');
  footerPara.addField(Field.createPageNumber());
  footerPara.addText(' of ');
  footerPara.addField(Field.createTotalPages());

  // Set header and footer
  doc.setHeader(header);
  doc.setFooter(footer);

  // Add content
  doc.createParagraph('Simple Header and Footer Example')
    .setStyle('Title')
    .setSpaceAfter(480);

  doc.createParagraph('Introduction')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  doc.createParagraph(
    'This document demonstrates a simple header and footer. The header contains the document title ' +
    'aligned to the right, and the footer contains centered page numbers in the format "Page X of Y".'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Add more content to show multiple pages
  for (let i = 1; i <= 3; i++) {
    doc.createParagraph(`Section ${i}`)
      .setStyle('Heading2')
      .setSpaceBefore(480)
      .setSpaceAfter(240);

    doc.createParagraph(
      'Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut ' +
      'labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco ' +
      'laboris nisi ut aliquip ex ea commodo consequat. '.repeat(3)
    )
      .setAlignment('justify')
      .setSpaceAfter(240);
  }

  // Save document
  const outputPath = path.join(outputDir, 'example1-simple-header-footer.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 2: Header with date and footer with filename
 */
async function example2_DateAndFilename() {
  console.log('Example 2: Header with date and footer with filename...');

  const doc = Document.create({
    properties: {
      title: 'Date and Filename Example',
      creator: 'DocXML Examples',
    },
  });

  // Create header with date
  const header = Header.createDefault();
  const headerPara = header.createParagraph();
  headerPara.setAlignment('right');
  headerPara.addField(Field.createDate('MMMM d, yyyy', { italic: true }));

  // Create footer with filename and page number
  const footer = Footer.createDefault();
  const footerPara = footer.createParagraph();
  footerPara.setAlignment('center');
  footerPara.addField(Field.createFilename(false, { size: 9 }));
  footerPara.addText(' - Page ', { size: 9 });
  footerPara.addField(Field.createPageNumber({ size: 9 }));

  // Set header and footer
  doc.setHeader(header);
  doc.setFooter(footer);

  // Add content
  doc.createParagraph('Document with Date and Filename')
    .setStyle('Title')
    .setSpaceAfter(480);

  doc.createParagraph(
    'This document shows how to use dynamic fields in headers and footers. The header displays ' +
    'the current date, and the footer shows the filename and page number.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Add sample content
  doc.createParagraph('Content Section')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  doc.createParagraph(
    'The date field updates automatically when the document is opened in Microsoft Word. ' +
    'The filename field shows the document name without the path. '.repeat(5)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Save document
  const outputPath = path.join(outputDir, 'example2-date-filename.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 3: Different first page header/footer
 */
async function example3_DifferentFirstPage() {
  console.log('Example 3: Different first page header/footer...');

  const doc = Document.create({
    properties: {
      title: 'Different First Page Example',
      creator: 'DocXML Examples',
    },
  });

  // Create first page header (title page - no header)
  const firstHeader = Header.createFirst();
  // Leave empty for title page

  // Create default header (for other pages)
  const header = Header.createDefault();
  const headerPara = header.createParagraph();
  headerPara.setAlignment('right');
  headerPara.addText('Company Report 2025', { bold: true, size: 10 });

  // Create first page footer (no page number on title page)
  const firstFooter = Footer.createFirst();
  const firstFooterPara = firstFooter.createParagraph();
  firstFooterPara.setAlignment('center');
  firstFooterPara.addText('Confidential', { italic: true, size: 9, color: '666666' });

  // Create default footer (with page numbers)
  const footer = Footer.createDefault();
  const footerPara = footer.createParagraph();
  footerPara.setAlignment('center');
  footerPara.addText('Page ');
  footerPara.addField(Field.createPageNumber({ size: 10 }));

  // Set headers and footers
  doc.setFirstPageHeader(firstHeader);
  doc.setHeader(header);
  doc.setFirstPageFooter(firstFooter);
  doc.setFooter(footer);

  // Add title page content
  doc.createParagraph('Annual Report')
    .setStyle('Title')
    .setSpaceBefore(2880) // 2 inches from top
    .setSpaceAfter(480);

  doc.createParagraph('2025')
    .setStyle('Subtitle')
    .setSpaceAfter(1440);

  doc.createParagraph('Company Name')
    .setAlignment('center')
    .addText('Company Name', { bold: true, size: 14 });

  // Add page break
  doc.createParagraph('')
    .setPageBreakBefore(true);

  // Add content pages
  doc.createParagraph('Executive Summary')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  doc.createParagraph(
    'This document demonstrates different headers and footers for the first page versus ' +
    'subsequent pages. The title page has no header and a simple "Confidential" footer, ' +
    'while other pages have the company name in the header and page numbers in the footer. '.repeat(3)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Key Findings')
    .setStyle('Heading2')
    .setSpaceAfter(240);

  doc.createParagraph(
    'Additional content appears on subsequent pages with the standard header and footer. '.repeat(4)
  )
    .setAlignment('justify');

  // Save document
  const outputPath = path.join(outputDir, 'example3-different-first-page.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 4: Header with table
 */
async function example4_HeaderWithTable() {
  console.log('Example 4: Header with table layout...');

  const doc = Document.create({
    properties: {
      title: 'Header with Table Example',
      creator: 'DocXML Examples',
    },
  });

  // Create header with 3-column table
  const header = Header.createDefault();
  const table = header.createTable(1, 3);

  // Left column - company name
  table.getCell(0, 0)?.createParagraph().addText('ACME Corp', { bold: true, size: 10 });

  // Center column - document title
  table.getCell(0, 1)
    ?.createParagraph()
    .setAlignment('center')
    .addText('Technical Report', { bold: true, size: 10 });

  // Right column - date
  table.getCell(0, 2)
    ?.createParagraph()
    .setAlignment('right')
    .addField(Field.createDate('MM/dd/yyyy', { size: 9 }));

  // Set table width to full page
  table.setWidth(9360); // Full page width
  table.setLayout('fixed');

  // Create footer with line and page number
  const footer = Footer.createDefault();

  // Add a horizontal line (using underline)
  const line = footer.createParagraph();
  line.addText('_'.repeat(80), { size: 8, color: 'CCCCCC' });

  // Add page number
  const footerPara = footer.createParagraph();
  footerPara.setAlignment('center');
  footerPara.setSpaceBefore(60);
  footerPara.addField(Field.createPageNumber({ size: 10 }));

  // Set header and footer
  doc.setHeader(header);
  doc.setFooter(footer);

  // Add content
  doc.createParagraph('Professional Header Layout')
    .setStyle('Title')
    .setSpaceAfter(480);

  doc.createParagraph(
    'This document uses a table in the header to create a professional three-column layout. ' +
    'The left column shows the company name, the center has the document title, and the ' +
    'right column displays the current date.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Benefits of Table Headers')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  doc.createParagraph(
    'Using tables in headers allows for precise control over column alignment and spacing. ' +
    'This is a common pattern in business documents and reports. '.repeat(4)
  )
    .setAlignment('justify');

  // Save document
  const outputPath = path.join(outputDir, 'example4-header-table.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 5: Multiple field types
 */
async function example5_MultipleFields() {
  console.log('Example 5: Document with multiple field types...');

  const doc = Document.create({
    properties: {
      title: 'Field Types Demo',
      creator: 'DocXML Examples',
      subject: 'Demonstrating Various Fields',
    },
  });

  // Create header with author and title
  const header = Header.createDefault();
  const headerPara = header.createParagraph();
  headerPara.setAlignment('right');
  headerPara.addField(Field.createAuthor({ size: 9 }));
  headerPara.addText(' - ', { size: 9 });
  headerPara.addField(Field.createTitle({ bold: true, size: 9 }));

  // Create footer with multiple fields
  const footer = Footer.createDefault();
  const footerLeft = footer.createParagraph();
  footerLeft.addText('Created: ', { size: 8 });
  footerLeft.addField(Field.createDate('MM/dd/yyyy', { size: 8 }));

  const footerCenter = footer.createParagraph();
  footerCenter.setAlignment('center');
  footerCenter.addField(Field.createPageNumber({ size: 10, bold: true }));

  // Set header and footer
  doc.setHeader(header);
  doc.setFooter(footer);

  // Add content explaining fields
  doc.createParagraph('Dynamic Field Types')
    .setStyle('Title')
    .setSpaceAfter(480);

  doc.createParagraph('Available Field Types')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  const fieldTypes = [
    'PAGE - Current page number',
    'NUMPAGES - Total number of pages',
    'DATE - Current date (with custom formatting)',
    'TIME - Current time',
    'AUTHOR - Document author',
    'TITLE - Document title',
    'FILENAME - Document filename',
    'SUBJECT - Document subject',
  ];

  for (const type of fieldTypes) {
    doc.createParagraph(`• ${type}`).setLeftIndent(360);
  }

  doc.createParagraph()
    .setSpaceBefore(480)
    .setAlignment('justify')
    .addText(
      'All these fields update automatically when the document is opened in Microsoft Word. ' +
      'Fields can be formatted with bold, italic, colors, and different font sizes just like regular text. '.repeat(2)
    );

  // Save document
  const outputPath = path.join(outputDir, 'example5-multiple-fields.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Main function to run all examples
 */
async function main() {
  console.log('Running Header/Footer Examples...\n');

  try {
    await example1_SimpleHeaderFooter();
    await example2_DateAndFilename();
    await example3_DifferentFirstPage();
    await example4_HeaderWithTable();
    await example5_MultipleFields();

    console.log('\n✓ All examples completed successfully!');
    console.log(`\nOutput files saved to: ${outputDir}`);
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
  example1_SimpleHeaderFooter,
  example2_DateAndFilename,
  example3_DifferentFirstPage,
  example4_HeaderWithTable,
  example5_MultipleFields,
};
