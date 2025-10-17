/**
 * Hyperlink Usage Examples
 *
 * This example demonstrates how to add hyperlinks to documents.
 * Includes external links, email links, and custom formatted links.
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
 * Example 1: Simple web links
 */
async function example1_SimpleWebLinks() {
  console.log('Example 1: Simple web links...');

  const doc = Document.create({
    properties: {
      title: 'Simple Web Links Example',
      creator: 'DocXML Examples',
    },
  });

  // Title
  doc.createParagraph('Web Links Demo')
    .setStyle('Title')
    .setSpaceAfter(480);

  // Introduction
  doc.createParagraph(
    'This document demonstrates how to create hyperlinks to external websites. ' +
    'Click any of the blue underlined links below to visit the website.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Add some web links
  doc.createParagraph('Common Websites')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  // Google link
  const para1 = doc.createParagraph();
  para1.addText('Search: ');
  para1.addHyperlink(Hyperlink.createWebLink('https://www.google.com', 'Google'));
  para1.setSpaceAfter(120);

  // GitHub link
  const para2 = doc.createParagraph();
  para2.addText('Code hosting: ');
  para2.addHyperlink(Hyperlink.createWebLink('https://github.com', 'GitHub'));
  para2.setSpaceAfter(120);

  // Stack Overflow link
  const para3 = doc.createParagraph();
  para3.addText('Q&A: ');
  para3.addHyperlink(Hyperlink.createWebLink('https://stackoverflow.com', 'Stack Overflow'));
  para3.setSpaceAfter(120);

  // Documentation link
  const para4 = doc.createParagraph();
  para4.addText('TypeScript docs: ');
  para4.addHyperlink(Hyperlink.createWebLink('https://www.typescriptlang.org/docs/', 'TypeScript Documentation'));

  // Save document
  const outputPath = path.join(outputDir, 'example1-simple-web-links.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 2: Email links
 */
async function example2_EmailLinks() {
  console.log('Example 2: Email links...');

  const doc = Document.create({
    properties: {
      title: 'Email Links Example',
      creator: 'DocXML Examples',
    },
  });

  // Title
  doc.createParagraph('Contact Information')
    .setStyle('Title')
    .setSpaceAfter(480);

  // Introduction
  doc.createParagraph(
    'This document demonstrates email links. Click any email address to open your default email client.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Contact section
  doc.createParagraph('Our Team')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  // Email links
  const para1 = doc.createParagraph();
  para1.addText('Sales: ');
  para1.addHyperlink(Hyperlink.createEmail('sales@example.com'));
  para1.setSpaceAfter(120);

  const para2 = doc.createParagraph();
  para2.addText('Support: ');
  para2.addHyperlink(Hyperlink.createEmail('support@example.com'));
  para2.setSpaceAfter(120);

  const para3 = doc.createParagraph();
  para3.addText('General inquiries: ');
  para3.addHyperlink(Hyperlink.createEmail('info@example.com', 'Contact Us'));

  // Save document
  const outputPath = path.join(outputDir, 'example2-email-links.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 3: Custom formatted links
 */
async function example3_CustomFormattedLinks() {
  console.log('Example 3: Custom formatted links...');

  const doc = Document.create({
    properties: {
      title: 'Custom Formatted Links Example',
      creator: 'DocXML Examples',
    },
  });

  // Title
  doc.createParagraph('Custom Link Styling')
    .setStyle('Title')
    .setSpaceAfter(480);

  // Introduction
  doc.createParagraph(
    'Links can be styled with custom colors, fonts, and formatting. ' +
    'The examples below show different link styles.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Red bold link
  doc.createParagraph('Styled Links')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  const para1 = doc.createParagraph();
  para1.addText('Red bold link: ');
  para1.addHyperlink(
    Hyperlink.createWebLink(
      'https://example.com',
      'Red Bold Link',
      { bold: true, color: 'FF0000' }
    )
  );
  para1.setSpaceAfter(120);

  // Green italic link
  const para2 = doc.createParagraph();
  para2.addText('Green italic link: ');
  para2.addHyperlink(
    Hyperlink.createWebLink(
      'https://example.com',
      'Green Italic Link',
      { italic: true, color: '008000' }
    )
  );
  para2.setSpaceAfter(120);

  // Large purple link
  const para3 = doc.createParagraph();
  para3.addText('Large purple link: ');
  para3.addHyperlink(
    Hyperlink.createWebLink(
      'https://example.com',
      'Large Purple Link',
      { size: 16, color: '800080', bold: true }
    )
  );
  para3.setSpaceAfter(120);

  // Highlighted link
  const para4 = doc.createParagraph();
  para4.addText('Highlighted link: ');
  para4.addHyperlink(
    Hyperlink.createWebLink(
      'https://example.com',
      'Highlighted Link',
      { highlight: 'yellow', bold: true }
    )
  );

  // Save document
  const outputPath = path.join(outputDir, 'example3-custom-formatted-links.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 4: Links in different contexts
 */
async function example4_LinksInContext() {
  console.log('Example 4: Links in different contexts...');

  const doc = Document.create({
    properties: {
      title: 'Links in Context Example',
      creator: 'DocXML Examples',
    },
  });

  // Title
  doc.createParagraph('Hyperlinks in Context')
    .setStyle('Title')
    .setSpaceAfter(480);

  // Paragraph with inline links
  const para1 = doc.createParagraph();
  para1.addText('You can embed links within sentences. For example, check out ');
  para1.addHyperlink(Hyperlink.createWebLink('https://github.com', 'GitHub'));
  para1.addText(' for code hosting, ');
  para1.addHyperlink(Hyperlink.createWebLink('https://stackoverflow.com', 'Stack Overflow'));
  para1.addText(' for Q&A, and ');
  para1.addHyperlink(Hyperlink.createWebLink('https://www.typescriptlang.org', 'TypeScript'));
  para1.addText(' for TypeScript documentation.');
  para1.setAlignment('justify');
  para1.setSpaceAfter(240);

  // Reference section
  doc.createParagraph('References')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  // Numbered list with links
  const ref1 = doc.createParagraph();
  ref1.addText('1. ');
  ref1.addHyperlink(Hyperlink.createWebLink('https://www.ecma-international.org/publications-and-standards/standards/ecma-376/', 'ECMA-376: Office Open XML File Formats'));
  ref1.setLeftIndent(360);
  ref1.setSpaceAfter(120);

  const ref2 = doc.createParagraph();
  ref2.addText('2. ');
  ref2.addHyperlink(Hyperlink.createWebLink('https://www.iso.org/standard/71691.html', 'ISO/IEC 29500: Information technology — Office Open XML formats'));
  ref2.setLeftIndent(360);
  ref2.setSpaceAfter(120);

  const ref3 = doc.createParagraph();
  ref3.addText('3. ');
  ref3.addHyperlink(Hyperlink.createWebLink('https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-wordprocessingml-document', 'Microsoft: Structure of a WordprocessingML document'));
  ref3.setLeftIndent(360);

  // Save document
  const outputPath = path.join(outputDir, 'example4-links-in-context.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 5: Internal links with bookmarks
 */
async function example5_InternalLinks() {
  console.log('Example 5: Internal links (bookmarks)...');

  const doc = Document.create({
    properties: {
      title: 'Internal Links Example',
      creator: 'DocXML Examples',
    },
  });

  // Title
  doc.createParagraph('Internal Document Links')
    .setStyle('Title')
    .setSpaceAfter(480);

  // Table of contents (manual)
  doc.createParagraph('Table of Contents')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  const toc1 = doc.createParagraph();
  toc1.addHyperlink(Hyperlink.createInternal('section1', 'Section 1: Introduction'));
  toc1.setLeftIndent(360);
  toc1.setSpaceAfter(120);

  const toc2 = doc.createParagraph();
  toc2.addHyperlink(Hyperlink.createInternal('section2', 'Section 2: Features'));
  toc2.setLeftIndent(360);
  toc2.setSpaceAfter(120);

  const toc3 = doc.createParagraph();
  toc3.addHyperlink(Hyperlink.createInternal('section3', 'Section 3: Conclusion'));
  toc3.setLeftIndent(360);
  toc3.setSpaceAfter(480);

  // Sections (bookmarks would need to be added in a future phase)
  // For now, just create the content
  doc.createParagraph('Section 1: Introduction')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  doc.createParagraph(
    'This is the introduction section. Internal links (also called bookmarks) allow you to ' +
    'link to specific sections within the same document. '.repeat(2)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('Section 2: Features')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  doc.createParagraph(
    'This section describes the features. Note: Full bookmark support requires additional ' +
    'implementation in the Document class. '.repeat(2)
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  doc.createParagraph('Section 3: Conclusion')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  doc.createParagraph(
    'This is the conclusion section. Internal links are useful for creating navigation within ' +
    'long documents. '.repeat(2)
  )
    .setAlignment('justify');

  // Save document
  const outputPath = path.join(outputDir, 'example5-internal-links.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 6: Validation and best practices
 *
 * This example demonstrates the validation rules enforced by DocXML
 * to prevent document corruption per ECMA-376 standards.
 */
async function example6_ValidationAndBestPractices() {
  console.log('Example 6: Validation and best practices...');

  const doc = Document.create({
    properties: {
      title: 'Hyperlink Validation Example',
      creator: 'DocXML Examples',
    },
  });

  // Title
  doc.createParagraph('Hyperlink Validation and Best Practices')
    .setStyle('Title')
    .setSpaceAfter(480);

  // Introduction
  doc.createParagraph(
    'DocXML enforces ECMA-376 compliance to prevent document corruption. ' +
    'External hyperlinks require relationship registration, which is ' +
    'automatically handled when using Document.save() or Document.toBuffer().'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  // Best Practice #1
  doc.createParagraph('✅ Best Practice: Use Document.save()')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  doc.createParagraph(
    'The recommended pattern is to use Document.save() or Document.toBuffer(). ' +
    'These methods automatically register all external hyperlinks with the ' +
    'relationship manager, ensuring valid OpenXML documents.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  const example1 = doc.createParagraph();
  example1.addText('Example: ');
  example1.addHyperlink(Hyperlink.createExternal('https://example.com', 'This link works correctly'));
  example1.addText(' because Document.save() handles relationships automatically.');
  example1.setSpaceAfter(480);

  // Validation Rule #1
  doc.createParagraph('🔒 Validation: External Links Require Relationship IDs')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  doc.createParagraph(
    'Per ECMA-376 Part 1 §17.16.22, external hyperlinks MUST have a relationship ID. ' +
    'DocXML throws an error if you attempt to generate XML for an external link without ' +
    'one, preventing silent document corruption.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  const validation1 = doc.createParagraph();
  validation1.addText('Attempting to manually call toXML() on an external hyperlink ');
  validation1.addText('without a relationship ID will throw: ', { italic: true });
  validation1.addText('"CRITICAL: External hyperlink to [URL] is missing relationship ID."',
    { color: 'FF0000', font: 'Courier New', size: 10 });
  validation1.setSpaceAfter(480);

  // Validation Rule #2
  doc.createParagraph('🔒 Validation: Empty Hyperlinks Rejected')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  doc.createParagraph(
    'Hyperlinks must have either a URL (external) or anchor (internal). ' +
    'DocXML rejects empty hyperlinks that have neither.'
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  // Validation Rule #3
  doc.createParagraph('⚠️ Warning: Hybrid Links (URL + Anchor)')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  doc.createParagraph(
    'Hyperlinks with both a URL and an anchor are ambiguous per ECMA-376. ' +
    'DocXML logs a warning when such links are created, though the URL takes precedence.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  doc.createParagraph(
    'Use Hyperlink.createExternal() for web links or Hyperlink.createInternal() ' +
    'for bookmark links to avoid this ambiguity.'
  )
    .setAlignment('justify')
    .setSpaceAfter(480);

  // Improved Text Fallback
  doc.createParagraph('✨ Improved Text Fallback (v0.3.0+)')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  doc.createParagraph(
    'When hyperlink text is empty, DocXML uses an improved fallback chain: ' +
    'text → url → anchor → "Link". This makes links more user-friendly.'
  )
    .setAlignment('justify')
    .setSpaceAfter(240);

  const fallback1 = doc.createParagraph();
  fallback1.addText('Example with empty text: ');
  // This will display "https://example.com" as the link text
  fallback1.addHyperlink(Hyperlink.createExternal('https://example.com', ''));
  fallback1.setSpaceAfter(480);

  // References
  doc.createParagraph('📚 References')
    .setStyle('Heading1')
    .setSpaceAfter(240);

  const ref1 = doc.createParagraph();
  ref1.addText('• ECMA-376 Part 1 §17.16.22: ');
  ref1.addHyperlink(
    Hyperlink.createExternal(
      'https://www.ecma-international.org/publications-and-standards/standards/ecma-376/',
      'Hyperlink Element Specification'
    )
  );
  ref1.setLeftIndent(360);
  ref1.setSpaceAfter(120);

  const ref2 = doc.createParagraph();
  ref2.addText('• DocXML Hyperlink Best Practices: See OPENXML_STRUCTURE_GUIDE.md');
  ref2.setLeftIndent(360);

  // Save document
  const outputPath = path.join(outputDir, 'example6-validation-best-practices.docx');
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Main function to run all examples
 */
async function main() {
  console.log('Running Hyperlink Examples...\n');

  try {
    await example1_SimpleWebLinks();
    await example2_EmailLinks();
    await example3_CustomFormattedLinks();
    await example4_LinksInContext();
    await example5_InternalLinks();
    await example6_ValidationAndBestPractices();

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
  example1_SimpleWebLinks,
  example2_EmailLinks,
  example3_CustomFormattedLinks,
  example4_LinksInContext,
  example5_InternalLinks,
  example6_ValidationAndBestPractices,
};
