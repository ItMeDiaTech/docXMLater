/**
 * Phase 3 Feature Demonstration
 * Showcases: Styles, Tables, Sections, and Numbering
 */

import { Document } from '../src';
import { Style } from '../src/formatting/Style';
import { PAGE_SIZES } from '../src/utils/units';

async function createPhase3Demo() {
  console.log('📝 Creating Phase 3 feature demonstration document...\n');

  // Create document
  const doc = Document.create();

  // ============= CUSTOM STYLES =============
  console.log('1️⃣ Adding custom styles...');

  // Create a custom title style
  const customTitleStyle = Style.create({
    styleId: 'CustomTitle',
    name: 'Custom Title',
    type: 'paragraph',
    basedOn: 'Normal',
    next: 'Normal',
    paragraphFormatting: {
      alignment: 'center',
      spacing: {
        before: 240,
        after: 360,
      },
    },
    runFormatting: {
      font: 'Georgia',
      size: 24,
      bold: true,
      color: '2E74B5',
    },
  });
  doc.addStyle(customTitleStyle);

  // Create a custom body style
  const customBodyStyle = Style.create({
    styleId: 'CustomBody',
    name: 'Custom Body',
    type: 'paragraph',
    basedOn: 'Normal',
    paragraphFormatting: {
      alignment: 'justify',
      indentation: {
        firstLine: 720, // 0.5 inch first line indent
      },
      spacing: {
        after: 120,
        line: 360,
        lineRule: 'auto',
      },
    },
    runFormatting: {
      font: 'Times New Roman',
      size: 12,
    },
  });
  doc.addStyle(customBodyStyle);

  // Create a code style
  const codeStyle = Style.create({
    styleId: 'CodeBlock',
    name: 'Code Block',
    type: 'paragraph',
    basedOn: 'Normal',
    paragraphFormatting: {
      spacing: {
        before: 120,
        after: 120,
      },
    },
    runFormatting: {
      font: 'Consolas',
      size: 10,
      color: '1F497D',
    },
  });
  doc.addStyle(codeStyle);

  // Add standard heading styles
  doc.addStyle(Style.createHeadingStyle(1));
  doc.addStyle(Style.createHeadingStyle(2));
  doc.addStyle(Style.createHeadingStyle(3));

  // ============= SECTION CONFIGURATION =============
  console.log('2️⃣ Configuring document section...');

  // Set up section with custom page size and margins
  doc.getSection()
    .setPageSize(PAGE_SIZES.LETTER.width, PAGE_SIZES.LETTER.height, 'portrait')
    .setMargins({
      top: 1440,    // 1 inch
      bottom: 1440,
      left: 1800,   // 1.25 inches
      right: 1800,
      header: 720,
      footer: 720,
    })
    .setPageNumbering(1, 'decimal');

  // ============= DOCUMENT CONTENT =============
  console.log('3️⃣ Adding document content...');

  // Title using custom style
  const title = doc.createParagraph('Phase 3 Features Demonstration');
  title.setStyle('CustomTitle');

  // Subtitle
  const subtitle = doc.createParagraph('Styles, Tables, Sections, and Numbering');
  subtitle.setStyle('Subtitle');
  subtitle.setAlignment('center');

  // Introduction
  doc.createParagraph(); // Empty paragraph for spacing
  const intro = doc.createParagraph(
    'This document demonstrates the Phase 3 features of the docXMLater library, ' +
    'including custom styles, advanced table formatting, section configuration, ' +
    'and multi-level numbering lists.'
  );
  intro.setStyle('CustomBody');

  // ============= NUMBERED LISTS =============
  console.log('4️⃣ Creating numbered and bullet lists...');

  // Heading for lists section
  const listsHeading = doc.createParagraph('Document Features');
  listsHeading.setStyle('Heading1');

  // Create a numbered list
  const numberedListId = doc.createNumberedList();

  const feature1 = doc.createParagraph('Custom Styles');
  feature1.setNumbering(numberedListId, 0);

  const feature2 = doc.createParagraph('Advanced Tables');
  feature2.setNumbering(numberedListId, 0);

  const feature3 = doc.createParagraph('Section Configuration');
  feature3.setNumbering(numberedListId, 0);

  const feature4 = doc.createParagraph('Multi-level Lists');
  feature4.setNumbering(numberedListId, 0);

  // Create a bullet list with sub-items
  const bulletListId = doc.createBulletList();

  doc.createParagraph(); // Space before bullet list
  const bulletHeading = doc.createParagraph('Implementation Details');
  bulletHeading.setStyle('Heading2');

  const bullet1 = doc.createParagraph('Style Management');
  bullet1.setNumbering(bulletListId, 0);

  const bullet1a = doc.createParagraph('Paragraph styles for consistent formatting');
  bullet1a.setNumbering(bulletListId, 1);

  const bullet1b = doc.createParagraph('Character styles for inline formatting');
  bullet1b.setNumbering(bulletListId, 1);

  const bullet2 = doc.createParagraph('Table Features');
  bullet2.setNumbering(bulletListId, 0);

  const bullet2a = doc.createParagraph('Cell merging and spanning');
  bullet2a.setNumbering(bulletListId, 1);

  const bullet2b = doc.createParagraph('Custom borders and shading');
  bullet2b.setNumbering(bulletListId, 1);

  const bullet2c = doc.createParagraph('Column width control');
  bullet2c.setNumbering(bulletListId, 1);

  // ============= TABLES =============
  console.log('5️⃣ Creating formatted tables...');

  doc.createParagraph(); // Space before table
  const tableHeading = doc.createParagraph('Feature Comparison Table');
  tableHeading.setStyle('Heading1');

  // Create a formatted table
  const table = doc.createTable(5, 4);

  // Configure table formatting
  table
    .setWidth(9000) // 6.25 inches
    .setAlignment('center')
    .setAllBorders({
      style: 'single',
      size: 4,
      color: '4472A8',
    })
    .setCellSpacing(0)
    .setColumnWidths([2250, 2250, 2250, 2250]); // Equal columns

  // Header row
  const headerRow = table.getRow(0);
  headerRow?.setHeader(true);
  headerRow?.setHeight(600); // Taller header

  // Fill header cells
  headerRow?.getCell(0)?.createParagraph('Feature');
  headerRow?.getCell(0)?.setShading({ fill: '4472A8' });
  headerRow?.getCell(1)?.createParagraph('Phase 1');
  headerRow?.getCell(1)?.setShading({ fill: '4472A8' });
  headerRow?.getCell(2)?.createParagraph('Phase 2');
  headerRow?.getCell(2)?.setShading({ fill: '4472A8' });
  headerRow?.getCell(3)?.createParagraph('Phase 3');
  headerRow?.getCell(3)?.setShading({ fill: '4472A8' });

  // Make header text white and bold
  for (let i = 0; i < 4; i++) {
    const cell = headerRow?.getCell(i);
    const para = cell?.getParagraphs()[0];
    if (para) {
      para.getRuns()[0]?.setBold(true).setColor('FFFFFF');
      para.setAlignment('center');
    }
  }

  // Data rows with alternating shading
  const data = [
    ['ZIP Handling', '✅ Complete', 'N/A', 'N/A'],
    ['Basic Elements', 'N/A', '✅ Complete', 'Enhanced'],
    ['Styles', 'N/A', 'N/A', '✅ Complete'],
    ['Tables', 'N/A', 'Basic', '✅ Advanced'],
  ];

  for (let row = 0; row < data.length; row++) {
    const tableRow = table.getRow(row + 1);
    const shading = row % 2 === 0 ? 'F2F2F2' : undefined;

    for (let col = 0; col < 4; col++) {
      const cell = tableRow?.getCell(col);
      if (cell) {
        cell.createParagraph(data[row]![col]!);
        if (shading) {
          cell.setShading({ fill: shading });
        }
        // Center align status columns
        if (col > 0) {
          cell.getParagraphs()[0]?.setAlignment('center');
        }
      }
    }
  }

  // ============= CODE BLOCK EXAMPLE =============
  console.log('6️⃣ Adding code example...');

  doc.createParagraph(); // Space
  const codeHeading = doc.createParagraph('Code Example');
  codeHeading.setStyle('Heading2');

  const codePara1 = doc.createParagraph('// Creating a custom style');
  codePara1.setStyle('CodeBlock');

  const codePara2 = doc.createParagraph('const style = Style.create({');
  codePara2.setStyle('CodeBlock');

  const codePara3 = doc.createParagraph('  styleId: "MyStyle",');
  codePara3.setStyle('CodeBlock');

  const codePara4 = doc.createParagraph('  name: "My Custom Style",');
  codePara4.setStyle('CodeBlock');

  const codePara5 = doc.createParagraph('  type: "paragraph"');
  codePara5.setStyle('CodeBlock');

  const codePara6 = doc.createParagraph('});');
  codePara6.setStyle('CodeBlock');

  // ============= QUARTERLY REPORT TABLE =============
  console.log('7️⃣ Creating quarterly report table...');

  doc.createParagraph();
  const mergeTableHeading = doc.createParagraph('Quarterly Report');
  mergeTableHeading.setStyle('Heading2');

  const complexTable = doc.createTable(4, 5);
  complexTable
    .setWidth(9000)
    .setAlignment('center')
    .setAllBorders({
      style: 'double',
      size: 6,
      color: '2E74B5',
    });

  // Title row
  const topRow = complexTable.getRow(0);
  topRow?.getCell(0)?.createParagraph('Quarterly Performance Report');
  topRow?.getCell(0)?.setShading({ fill: '2E74B5' });
  topRow?.getCell(0)?.setColumnSpan(5); // Span across all columns
  topRow?.getCell(0)?.getParagraphs()[0]?.setAlignment('center');
  topRow?.getCell(0)?.getParagraphs()[0]?.getRuns()[0]?.setBold(true).setColor('FFFFFF').setSize(14);

  // Hide the other cells in the merged row
  for (let i = 1; i < 5; i++) {
    topRow?.getCell(i)?.setWidth(0);
  }

  // Quarter headers
  const quarterRow = complexTable.getRow(1);
  quarterRow?.getCell(0)?.createParagraph('Metric');
  quarterRow?.getCell(0)?.setShading({ fill: 'E7E6E6' });
  quarterRow?.getCell(1)?.createParagraph('Q1');
  quarterRow?.getCell(1)?.setShading({ fill: 'E7E6E6' });
  quarterRow?.getCell(2)?.createParagraph('Q2');
  quarterRow?.getCell(2)?.setShading({ fill: 'E7E6E6' });
  quarterRow?.getCell(3)?.createParagraph('Q3');
  quarterRow?.getCell(3)?.setShading({ fill: 'E7E6E6' });
  quarterRow?.getCell(4)?.createParagraph('Q4');
  quarterRow?.getCell(4)?.setShading({ fill: 'E7E6E6' });

  // Center align all headers
  for (let i = 0; i < 5; i++) {
    quarterRow?.getCell(i)?.getParagraphs()[0]?.setAlignment('center');
    quarterRow?.getCell(i)?.getParagraphs()[0]?.getRuns()[0]?.setBold(true);
  }

  // Sales data
  const dataRow1 = complexTable.getRow(2);
  dataRow1?.getCell(0)?.createParagraph('Sales');
  dataRow1?.getCell(0)?.setVerticalAlignment('center');
  dataRow1?.getCell(1)?.createParagraph('$1.2M');
  dataRow1?.getCell(2)?.createParagraph('$1.5M');
  dataRow1?.getCell(3)?.createParagraph('$1.8M');
  dataRow1?.getCell(4)?.createParagraph('$2.1M');

  // Growth percentage data
  const dataRow2 = complexTable.getRow(3);
  dataRow2?.getCell(0)?.createParagraph('Growth');
  dataRow2?.getCell(1)?.createParagraph('+15%');
  dataRow2?.getCell(2)?.createParagraph('+25%');
  dataRow2?.getCell(3)?.createParagraph('+20%');
  dataRow2?.getCell(4)?.createParagraph('+17%');

  // ============= MULTI-COLUMN SECTION =============
  console.log('8️⃣ Demonstrating section properties...');

  doc.createParagraph();
  const sectionHeading = doc.createParagraph('Section Configuration');
  sectionHeading.setStyle('Heading1');

  const sectionInfo = doc.createParagraph(
    'This document uses Letter size paper (8.5" × 11") with 1.25" left and right margins, ' +
    '1" top and bottom margins. The page numbering starts at 1 using decimal format. ' +
    'Headers and footers are configured with 0.5" spacing from the page edge.'
  );
  sectionInfo.setStyle('CustomBody');

  // ============= CONCLUSION =============
  console.log('9️⃣ Adding conclusion...');

  doc.createParagraph();
  const conclusionHeading = doc.createParagraph('Summary');
  conclusionHeading.setStyle('Heading1');

  const conclusion = doc.createParagraph(
    'Phase 3 of the docXMLater project successfully implements advanced document formatting capabilities. ' +
    'The library now supports custom styles for consistent formatting, complex table layouts with cell ' +
    'merging and spanning, comprehensive section configuration, and multi-level numbering systems. ' +
    'These features enable the creation of professional, well-structured documents programmatically.'
  );
  conclusion.setStyle('CustomBody');

  // Final stats
  doc.createParagraph();
  const statsHeading = doc.createParagraph('Implementation Statistics');
  statsHeading.setStyle('Heading2');

  const statsTable = doc.createTable(3, 2);
  statsTable.setWidth(6000).setAlignment('center');

  statsTable.getRow(0)?.getCell(0)?.createParagraph('Metric');
  statsTable.getRow(0)?.getCell(0)?.setShading({ fill: 'E7E6E6' });
  statsTable.getRow(0)?.getCell(1)?.createParagraph('Value');
  statsTable.getRow(0)?.getCell(1)?.setShading({ fill: 'E7E6E6' });

  statsTable.getRow(1)?.getCell(0)?.createParagraph('Total Tests');
  statsTable.getRow(1)?.getCell(1)?.createParagraph('226+');

  statsTable.getRow(2)?.getCell(0)?.createParagraph('Source Files');
  statsTable.getRow(2)?.getCell(1)?.createParagraph('40+');

  // Save document
  console.log('\n💾 Saving document...');
  await doc.save('examples/phase3-demo.docx');
  console.log('✅ Document saved as phase3-demo.docx');

  // Display statistics
  console.log('\n📊 Document Statistics:');
  console.log(`  - Paragraphs: ${doc.getParagraphCount()}`);
  console.log(`  - Tables: ${doc.getTableCount()}`);
  console.log(`  - Custom Styles: ${doc.getStyles().length}`);
  console.log(`  - Word Count: ${doc.getWordCount()}`);
}

// Run the demo
createPhase3Demo().catch(console.error);