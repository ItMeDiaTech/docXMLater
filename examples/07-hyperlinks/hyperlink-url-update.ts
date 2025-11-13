/**
 * Hyperlink URL Update Example
 *
 * Demonstrates how to update hyperlink URLs in existing documents.
 * This is useful for batch URL updates, migrating links to new domains,
 * or fixing broken/outdated links.
 */

import { Document } from '../../src';
import * as path from 'path';
import * as fs from 'fs';

// Ensure output directory exists
const outputDir = __dirname;
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

/**
 * Example: Update hyperlink URLs in a loaded document
 */
async function updateHyperlinkUrlsExample() {
  console.log('\n=== Hyperlink URL Update Example ===\n');

  // First, create a sample document with hyperlinks
  console.log('Step 1: Creating sample document with hyperlinks...');
  const doc = Document.create({
    properties: {
      title: 'Document with Links',
      creator: 'DocXML Examples',
    },
  });

  // Add title
  doc.createParagraph('Document with Hyperlinks')
    .setStyle('Title')
    .setSpaceAfter(480);

  // Add paragraphs with various hyperlinks
  const para1 = doc.createParagraph();
  para1.addText('Check out our ');
  para1.addHyperlink(
    Hyperlink.createExternal('https://old-website.com', 'old website')
  );
  para1.addText(' for more information.');
  para1.setSpaceAfter(240);

  const para2 = doc.createParagraph();
  para2.addText('Visit the ');
  para2.addHyperlink(
    Hyperlink.createExternal('https://old-docs.com/guide', 'documentation')
  );
  para2.addText(' to learn more.');
  para2.setSpaceAfter(240);

  const para3 = doc.createParagraph();
  para3.addText('Contact us at ');
  para3.addHyperlink(
    Hyperlink.createEmail('support@old-domain.com', 'support@old-domain.com')
  );
  para3.setSpaceAfter(240);

  // Add a link that won't be updated (not in map)
  const para4 = doc.createParagraph();
  para4.addText('External reference: ');
  para4.addHyperlink(
    Hyperlink.createExternal('https://example.com', 'Example Site')
  );

  // Save original document
  const originalPath = path.join(outputDir, 'original-with-links.docx');
  await doc.save(originalPath);
  console.log(`✓ Original document saved: ${originalPath}`);

  // Step 2: Load the document and update URLs
  console.log('\nStep 2: Loading document and updating URLs...');
  const loadedDoc = await Document.load(originalPath);

  // Define URL mappings (old URL → new URL)
  const urlMap = new Map([
    ['https://old-website.com', 'https://new-website.com'],
    ['https://old-docs.com/guide', 'https://new-docs.com/guide'],
    ['mailto:support@old-domain.com', 'mailto:support@new-domain.com'],
  ]);

  // Update hyperlink URLs
  const updatedCount = loadedDoc.updateHyperlinkUrls(urlMap);
  console.log(`✓ Updated ${updatedCount} hyperlink(s)`);

  // Display changes
  console.log('\nURL Changes:');
  urlMap.forEach((newUrl, oldUrl) => {
    console.log(`  ${oldUrl}\n    → ${newUrl}`);
  });

  // Save updated document
  const updatedPath = path.join(outputDir, 'updated-with-new-links.docx');
  await loadedDoc.save(updatedPath);
  console.log(`\n✓ Updated document saved: ${updatedPath}`);

  // Step 3: Verify the updates
  console.log('\nStep 3: Verifying updates...');
  const verifyDoc = await Document.load(updatedPath);
  const paragraphs = verifyDoc.getParagraphs();

  let verifiedCount = 0;
  for (const para of paragraphs) {
    for (const content of para.getContent()) {
      if (content instanceof Hyperlink && content.isExternal()) {
        const url = content.getUrl();
        if (url) {
          console.log(`  ✓ Hyperlink: "${content.getText()}" → ${url}`);
          verifiedCount++;
        }
      }
    }
  }

  console.log(`\n✓ Verified ${verifiedCount} hyperlink(s) in updated document`);
  console.log('\n=== Example Complete ===\n');
}

/**
 * Example: Batch update multiple documents
 */
async function batchUpdateExample() {
  console.log('\n=== Batch Update Example ===\n');

  // Common URL mapping for all documents
  const urlMap = new Map([
    ['https://old-company.com', 'https://new-company.com'],
    ['https://old-blog.com', 'https://new-blog.com'],
    ['mailto:info@old-company.com', 'mailto:info@new-company.com'],
  ]);

  // Simulate multiple documents (in practice, you'd load from different files)
  const documents = [
    { name: 'Document 1', links: ['https://old-company.com'] },
    { name: 'Document 2', links: ['https://old-blog.com'] },
    { name: 'Document 3', links: ['https://old-company.com', 'mailto:info@old-company.com'] },
  ];

  console.log('Processing multiple documents with shared URL mapping...\n');

  for (const docInfo of documents) {
    // Create sample document
    const doc = Document.create();
    doc.createParagraph(docInfo.name).setStyle('Heading1');

    // Add links
    for (const url of docInfo.links) {
      const para = doc.createParagraph();
      para.addHyperlink(Hyperlink.createExternal(url, url));
    }

    // Update URLs
    const updated = doc.updateHyperlinkUrls(urlMap);
    console.log(`  ${docInfo.name}: Updated ${updated} link(s)`);

    // Save (in practice, you'd save to original location)
    const filename = `batch-${docInfo.name.toLowerCase().replace(/\s+/g, '-')}.docx`;
    await doc.save(path.join(outputDir, filename));
  }

  console.log(`\n✓ Batch update complete`);
  console.log('\n=== Example Complete ===\n');
}

/**
 * Example: Advanced usage with validation
 */
async function advancedUpdateExample() {
  console.log('\n=== Advanced Update Example ===\n');

  // Create document with various hyperlink types
  const doc = Document.create();

  doc.createParagraph('Mixed Hyperlinks Example')
    .setStyle('Title')
    .setSpaceAfter(480);

  // External web links
  const para1 = doc.createParagraph();
  para1.addHyperlink(Hyperlink.createExternal('https://update-me.com', 'Will be updated'));
  para1.setSpaceAfter(120);

  // Internal bookmark link (should NOT be updated)
  const para2 = doc.createParagraph();
  para2.addHyperlink(Hyperlink.createInternal('Section1', 'Internal link (not updated)'));
  para2.setSpaceAfter(120);

  // Another external link
  const para3 = doc.createParagraph();
  para3.addHyperlink(Hyperlink.createExternal('https://also-update.com', 'Also updated'));

  // Define URL mapping
  const urlMap = new Map([
    ['https://update-me.com', 'https://updated.com'],
    ['https://also-update.com', 'https://also-updated.com'],
  ]);

  console.log('Document structure:');
  console.log('  - 2 external hyperlinks (will be updated)');
  console.log('  - 1 internal hyperlink (will be skipped)');

  // Update URLs
  const updated = doc.updateHyperlinkUrls(urlMap);
  console.log(`\n✓ Updated ${updated} hyperlink(s) (internal links skipped as expected)`);

  // Save
  const outputPath = path.join(outputDir, 'advanced-update-example.docx');
  await doc.save(outputPath);
  console.log(`✓ Saved to ${outputPath}`);

  console.log('\n=== Example Complete ===\n');
}

// Import Hyperlink class
import { Hyperlink } from '../../src/elements/Hyperlink';

/**
 * Main function to run all examples
 */
async function main() {
  try {
    await updateHyperlinkUrlsExample();
    await batchUpdateExample();
    await advancedUpdateExample();

    console.log('\n✅ All hyperlink URL update examples completed successfully!');
    console.log(`\nOutput files saved to: ${outputDir}`);
  } catch (error) {
    console.error('\n❌ Error running examples:', error);
    process.exit(1);
  }
}

// Run examples if executed directly
if (require.main === module) {
  main();
}

export {
  updateHyperlinkUrlsExample,
  batchUpdateExample,
  advancedUpdateExample,
};
