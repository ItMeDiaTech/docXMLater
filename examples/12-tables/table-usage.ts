/**
 * Examples showing Table usage
 */

import { Document } from '../src';

/**
 * Example 1: Simple table
 */
async function example1SimpleTable() {
  console.log('\n=== Example 1: Simple Table ===');

  const doc = Document.create();

  doc.createParagraph('Simple 3x3 Table:').setSpaceBefore(240);

  // Create a simple 3x3 table
  const table = doc.createTable(3, 3);

  // Populate cells
  for (let row = 0; row < 3; row++) {
    for (let col = 0; col < 3; col++) {
      const cell = table.getCell(row, col);
      if (cell) {
        cell.createParagraph(`Row ${row + 1}, Col ${col + 1}`);
      }
    }
  }

  await doc.save('example1-simple-table.docx');
  console.log('✓ Created example1-simple-table.docx');
}

/**
 * Example 2: Table with borders
 */
async function example2TableWithBorders() {
  console.log('\n=== Example 2: Table with Borders ===');

  const doc = Document.create();

  doc.createParagraph('Table with Borders:').setSpaceBefore(240);

  const table = doc.createTable(3, 3);

  // Set table borders
  table.setAllBorders({
    style: 'single',
    size: 8,
    color: '000000',
  });

  // Populate cells
  for (let row = 0; row < 3; row++) {
    for (let col = 0; col < 3; col++) {
      table.getCell(row, col)?.createParagraph(`Cell ${row},${col}`);
    }
  }

  await doc.save('example2-table-borders.docx');
  console.log('✓ Created example2-table-borders.docx');
}

/**
 * Example 3: Table with header row and shading
 */
async function example3TableWithHeader() {
  console.log('\n=== Example 3: Table with Header ===');

  const doc = Document.create();

  doc.createParagraph('Employee Table:').addText(' with header and shading', { bold: true });

  const table = doc.createTable(4, 3);
  table.setAllBorders({ style: 'single', size: 6, color: '000000' });

  // Header row
  const headerRow = table.getRow(0);
  if (headerRow) {
    headerRow.setHeader(true);

    const headers = ['Name', 'Department', 'Salary'];
    headers.forEach((header, idx) => {
      const cell = headerRow.getCell(idx);
      if (cell) {
        cell.setShading({ fill: '4472C4' });
        cell.createParagraph().addText(header, { bold: true, color: 'FFFFFF' });
      }
    });
  }

  // Data rows
  const data = [
    ['John Doe', 'Engineering', '$95,000'],
    ['Jane Smith', 'Marketing', '$85,000'],
    ['Bob Johnson', 'Sales', '$75,000'],
  ];

  data.forEach((rowData, rowIdx) => {
    rowData.forEach((cellData, colIdx) => {
      table.getCell(rowIdx + 1, colIdx)?.createParagraph(cellData);
    });
  });

  await doc.save('example3-table-header.docx');
  console.log('✓ Created example3-table-header.docx');
}

/**
 * Example 4: Table with cell merging and formatting
 */
async function example4AdvancedTable() {
  console.log('\n=== Example 4: Advanced Table ===');

  const doc = Document.create();

  doc.createParagraph('Advanced Table Formatting:');

  const table = doc.createTable(4, 4);
  table.setAllBorders({ style: 'single', size: 8, color: '333333' });
  table.setWidth(8640); // Full page width (~6 inches)

  // Title cell spanning all columns
  const titleCell = table.getCell(0, 0);
  if (titleCell) {
    titleCell.setColumnSpan(4);
    titleCell.setShading({ fill: '2E75B6' });
    titleCell.setVerticalAlignment('center');
    const para = titleCell.createParagraph();
    para.setAlignment('center');
    para.addText('Quarterly Report', { bold: true, size: 16, color: 'FFFFFF' });
  }

  // Headers
  const headers = ['Q1', 'Q2', 'Q3', 'Q4'];
  headers.forEach((header, idx) => {
    const cell = table.getCell(1, idx);
    if (cell) {
      cell.setShading({ fill: 'D9E1F2' });
      cell.createParagraph().addText(header, { bold: true }).setAlignment('center');
    }
  });

  // Data rows
  const quarters = ['$50K', '$62K', '$58K', '$71K'];
  quarters.forEach((value, idx) => {
    table.getCell(2, idx)?.createParagraph(value).setAlignment('center');
  });

  // Total row
  const totalLabel = table.getCell(3, 0);
  if (totalLabel) {
    totalLabel.setColumnSpan(3);
    totalLabel.setShading({ fill: 'FFF2CC' });
    totalLabel.createParagraph().addText('Total:', { bold: true }).setAlignment('right');
  }

  const totalValue = table.getCell(3, 3);
  if (totalValue) {
    totalValue.setShading({ fill: 'FFF2CC' });
    totalValue.createParagraph().addText('$241K', { bold: true, color: '0070C0' }).setAlignment('center');
  }

  await doc.save('example4-advanced-table.docx');
  console.log('✓ Created example4-advanced-table.docx');
}

/**
 * Example 5: Mixed content (paragraphs and tables)
 */
async function example5MixedContent() {
  console.log('\n=== Example 5: Mixed Content ===');

  const doc = Document.create({
    properties: {
      title: 'Sales Report',
      creator: 'DocXML',
    },
  });

  // Title
  doc.createParagraph().setAlignment('center').addText('Monthly Sales Report', { bold: true, size: 18 });

  doc.createParagraph(); // Empty line

  // Introduction
  doc.createParagraph('This report summarizes the sales performance for the current month.');

  doc.createParagraph(); // Empty line

  // First table
  doc.createParagraph().addText('Sales by Region', { bold: true, size: 14 });

  const regionTable = doc.createTable(4, 2);
  regionTable.setAllBorders({ style: 'single', size: 6, color: '000000' });

  // Headers
  regionTable.getCell(0, 0)?.createParagraph().addText('Region', { bold: true });
  regionTable.getCell(0, 1)?.createParagraph().addText('Sales', { bold: true });

  // Data
  const regions = [['North', '$125K'], ['South', '$98K'], ['East', '$112K']];
  regions.forEach((row, idx) => {
    regionTable.getCell(idx + 1, 0)?.createParagraph(row[0]);
    regionTable.getCell(idx + 1, 1)?.createParagraph(row[1]);
  });

  doc.createParagraph(); // Empty line

  // Commentary
  doc.createParagraph('The North region continues to lead in sales performance.');

  doc.createParagraph(); // Empty line

  // Second table
  doc.createParagraph().addText('Top Products', { bold: true, size: 14 });

  const productTable = doc.createTable(4, 2);
  productTable.setAllBorders({ style: 'single', size: 6, color: '000000' });

  // Headers
  productTable.getCell(0, 0)?.createParagraph().addText('Product', { bold: true });
  productTable.getCell(0, 1)?.createParagraph().addText('Units Sold', { bold: true });

  // Data
  const products = [['Widget A', '1,250'], ['Widget B', '980'], ['Widget C', '1,100']];
  products.forEach((row, idx) => {
    productTable.getCell(idx + 1, 0)?.createParagraph(row[0]);
    productTable.getCell(idx + 1, 1)?.createParagraph(row[1]);
  });

  await doc.save('example5-mixed-content.docx');
  console.log('✓ Created example5-mixed-content.docx');
}

// Run all examples
async function runExamples() {
  console.log('=== DocXML Table Examples ===');

  try {
    await example1SimpleTable();
    await example2TableWithBorders();
    await example3TableWithHeader();
    await example4AdvancedTable();
    await example5MixedContent();

    console.log('\n=== All examples completed successfully! ===');
    console.log('\nGenerated files:');
    console.log('  - example1-simple-table.docx');
    console.log('  - example2-table-borders.docx');
    console.log('  - example3-table-header.docx');
    console.log('  - example4-advanced-table.docx');
    console.log('  - example5-mixed-content.docx');
  } catch (error) {
    console.error('Error running examples:', error);
    process.exit(1);
  }
}

// Run if executed directly
if (require.main === module) {
  runExamples();
}
