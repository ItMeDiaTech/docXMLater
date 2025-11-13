/**
 * Quick test for styles functionality
 */

import { Document } from '../src';

async function testStyles() {
  console.log('Testing styles functionality...\n');

  const doc = Document.create({ properties: { title: 'Styles Test' } });

  // Test built-in styles
  console.log('✓ Document created with built-in styles');

  // Check if Normal style exists
  const normalStyle = doc.getStyle('Normal');
  console.log(`✓ Normal style exists: ${normalStyle ? 'Yes' : 'No'}`);

  // Check if Heading1 exists
  const heading1 = doc.getStyle('Heading1');
  console.log(`✓ Heading1 style exists: ${heading1 ? 'Yes' : 'No'}`);

  // Add paragraphs with different styles
  doc.createParagraph('Document Title').setStyle('Title');
  doc.createParagraph('This is a subtitle').setStyle('Subtitle');
  doc.createParagraph();
  doc.createParagraph('Chapter 1').setStyle('Heading1');
  doc.createParagraph('This is normal body text.').setStyle('Normal');
  doc.createParagraph();
  doc.createParagraph('Section 1.1').setStyle('Heading2');
  doc.createParagraph('More body text here with the Normal style.');

  console.log(`✓ Created ${doc.getParagraphCount()} paragraphs with styles`);

  // Get styles manager
  const stylesManager = doc.getStylesManager();
  console.log(`✓ StylesManager has ${stylesManager.getStyleCount()} styles`);

  // Save document
  await doc.save('styles-test.docx');
  console.log('\n✓ Created styles-test.docx');
  console.log('  Open in Microsoft Word to verify styles are applied correctly!');
}

testStyles().catch(console.error);
