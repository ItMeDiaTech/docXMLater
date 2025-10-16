/**
 * Table of Contents Usage Examples
 *
 * Demonstrates how to add a table of contents (TOC) to documents.
 * Word will automatically populate the TOC based on heading styles.
 */

import { Document, TableOfContents } from '../../src';
import * as path from 'path';
import * as fs from 'fs';

// Create output directory if it doesn't exist
const outputDir = __dirname;
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

/**
 * Example 1: Simple TOC with standard settings
 */
async function example1_SimpleTOC() {
  console.log('Example 1: Simple Table of Contents...');

  const doc = Document.create({
    properties: {
      title: 'Simple TOC Example',
      creator: 'DocXML Examples',
    },
  });

  // Add standard TOC (includes Heading1-3)
  doc.createTableOfContents();

  // Add content with headings
  doc.createParagraph('Chapter 1: Introduction')
    .setStyle('Heading1');

  doc.createParagraph(
    'This chapter introduces the main concepts. The table of contents above will ' +
    'automatically populate with all the headings in this document when you open it ' +
    'in Microsoft Word and update the field (right-click on TOC → Update Field).'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Section 1.1: Background')
    .setStyle('Heading2');

  doc.createParagraph(
    'Background information goes here. '.repeat(10)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Section 1.2: Objectives')
    .setStyle('Heading2');

  doc.createParagraph(
    'Project objectives are listed here. '.repeat(10)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('Chapter 2: Methodology')
    .setStyle('Heading1');

  doc.createParagraph(
    'This chapter describes the methodology. '.repeat(15)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Section 2.1: Approach')
    .setStyle('Heading2');

  doc.createParagraph(
    'The approach is detailed here. '.repeat(12)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Subsection 2.1.1: Methods')
    .setStyle('Heading3');

  doc.createParagraph(
    'Specific methods are described. '.repeat(10)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('Chapter 3: Results')
    .setStyle('Heading1');

  doc.createParagraph(
    'Results and findings are presented here. '.repeat(15)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('Chapter 4: Conclusion')
    .setStyle('Heading1');

  doc.createParagraph(
    'Final conclusions and recommendations. '.repeat(10)
  )
    .setAlignment('justify');

  // Save document
  const outputPath = path.join(outputDir, 'example1-simple-toc.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log('  Note: Open in Word and right-click the TOC to update it!');
}

/**
 * Example 2: Detailed TOC with 4 levels
 */
async function example2_DetailedTOC() {
  console.log('Example 2: Detailed TOC with 4 levels...');

  const doc = Document.create({
    properties: {
      title: 'Detailed TOC Example',
      creator: 'DocXML Examples',
    },
  });

  // Create detailed TOC with 4 levels
  const toc = TableOfContents.createDetailed('Table of Contents');
  doc.addTableOfContents(toc);

  // Add multi-level content
  doc.createParagraph('Part I: Fundamentals')
    .setStyle('Heading1');

  doc.createParagraph('Chapter 1: Basic Concepts')
    .setStyle('Heading2');

  doc.createParagraph('Section 1.1: Definitions')
    .setStyle('Heading3');

  doc.createParagraph('Subsection 1.1.1: Core Terms')
    .setStyle('Heading4');

  doc.createParagraph(
    'Content for core terms section. '.repeat(8)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Subsection 1.1.2: Related Concepts')
    .setStyle('Heading4');

  doc.createParagraph(
    'Content for related concepts. '.repeat(8)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Section 1.2: Principles')
    .setStyle('Heading3');

  doc.createParagraph(
    'Content about principles. '.repeat(10)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Chapter 2: Advanced Topics')
    .setStyle('Heading2');

  doc.createParagraph('Section 2.1: Complex Scenarios')
    .setStyle('Heading3');

  doc.createParagraph(
    'Content about complex scenarios. '.repeat(12)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('Part II: Applications')
    .setStyle('Heading1');

  doc.createParagraph('Chapter 3: Practical Use Cases')
    .setStyle('Heading2');

  doc.createParagraph(
    'Applications and use cases. '.repeat(15)
  )
    .setAlignment('justify');

  // Save document
  const outputPath = path.join(outputDir, 'example2-detailed-toc.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 3: Custom TOC properties
 */
async function example3_CustomTOC() {
  console.log('Example 3: Custom TOC with properties...');

  const doc = Document.create({
    properties: {
      title: 'Custom TOC Example',
      creator: 'DocXML Examples',
    },
  });

  // Create TOC with custom properties
  const toc = new TableOfContents({
    title: 'Document Contents',
    levels: 2, // Only include Heading1 and Heading2
    showPageNumbers: true,
    rightAlignPageNumbers: true,
    tabLeader: 'dot', // Use dots as tab leader
  });
  doc.addTableOfContents(toc);

  // Add content
  doc.createParagraph('Executive Summary')
    .setStyle('Heading1');

  doc.createParagraph(
    'High-level overview of the document. '.repeat(12)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Key Findings')
    .setStyle('Heading2');

  doc.createParagraph(
    'Important findings from the research. '.repeat(10)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Recommendations')
    .setStyle('Heading2');

  doc.createParagraph(
    'Actionable recommendations. '.repeat(10)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('Detailed Analysis')
    .setStyle('Heading1');

  doc.createParagraph(
    'In-depth analysis of the data. '.repeat(15)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Methodology')
    .setStyle('Heading2');

  doc.createParagraph(
    'Research methodology details. '.repeat(10)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Note: Heading3 won't appear in TOC since we set levels=2
  doc.createParagraph('Data Collection (not in TOC)')
    .setStyle('Heading3');

  doc.createParagraph(
    'This Heading3 will not appear in the TOC since we limited it to 2 levels. '.repeat(8)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('Conclusion')
    .setStyle('Heading1');

  doc.createParagraph(
    'Final thoughts and next steps. '.repeat(10)
  )
    .setAlignment('justify');

  // Save document
  const outputPath = path.join(outputDir, 'example3-custom-toc.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 4: Hyperlinked TOC (for web documents)
 */
async function example4_HyperlinkedTOC() {
  console.log('Example 4: Hyperlinked TOC...');

  const doc = Document.create({
    properties: {
      title: 'Hyperlinked TOC Example',
      creator: 'DocXML Examples',
    },
  });

  // Create hyperlinked TOC (entries are clickable, no page numbers)
  const toc = TableOfContents.createHyperlinked('Contents');
  doc.addTableOfContents(toc);

  // Add content
  doc.createParagraph('Welcome')
    .setStyle('Heading1');

  doc.createParagraph(
    'This document uses a hyperlinked table of contents. Click on any entry in the ' +
    'TOC to jump directly to that section. This is especially useful for documents ' +
    'that will be viewed on screen rather than printed. '.repeat(2)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('Getting Started')
    .setStyle('Heading1');

  doc.createParagraph('Installation')
    .setStyle('Heading2');

  doc.createParagraph(
    'Installation instructions go here. '.repeat(10)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Configuration')
    .setStyle('Heading2');

  doc.createParagraph(
    'Configuration details. '.repeat(10)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('Usage Guide')
    .setStyle('Heading1');

  doc.createParagraph('Basic Features')
    .setStyle('Heading2');

  doc.createParagraph(
    'Basic feature documentation. '.repeat(10)
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph('Advanced Features')
    .setStyle('Heading2');

  doc.createParagraph(
    'Advanced feature documentation. '.repeat(10)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('Troubleshooting')
    .setStyle('Heading1');

  doc.createParagraph(
    'Common issues and solutions. '.repeat(12)
  )
    .setAlignment('justify');

  // Save document
  const outputPath = path.join(outputDir, 'example4-hyperlinked-toc.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Main function to run all examples
 */
async function main() {
  console.log('Running Table of Contents Examples...\n');

  try {
    await example1_SimpleTOC();
    await example2_DetailedTOC();
    await example3_CustomTOC();
    await example4_HyperlinkedTOC();

    console.log('\n✓ All examples completed successfully!');
    console.log(`\nOutput files saved to: ${outputDir}`);
    console.log('\n📝 Important: Open the documents in Microsoft Word and right-click');
    console.log('   the TOC, then select "Update Field" to populate it with headings!');
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
  example1_SimpleTOC,
  example2_DetailedTOC,
  example3_CustomTOC,
  example4_HyperlinkedTOC,
};
