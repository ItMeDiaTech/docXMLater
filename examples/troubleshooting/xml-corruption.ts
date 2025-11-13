/**
 * XML Corruption Example - Common Mistake and How to Fix It
 *
 * This example demonstrates the most common mistake users make with docXMLater:
 * passing XML-like strings to text methods instead of using the API properly.
 */

import { Document, Paragraph, Run, detectCorruptionInDocument } from '../../src/index';
import * as path from 'path';

async function demonstrateXmlCorruption() {
  console.log('='.repeat(80));
  console.log('XML Corruption in Text - Common Mistake Demo');
  console.log('='.repeat(80));
  console.log();

  // SECTION 1: The Wrong Way (Causes Corruption)
  console.log('SECTION 1: The WRONG Way (Causes Corruption)');
  console.log('-'.repeat(80));

  const wrongDoc = Document.create();
  const wrongPara = wrongDoc.createParagraph();

  // MISTAKE: Passing XML tags as text
  wrongPara.addText('Important Information<w:t xml:space="preserve">1</w:t>');

  console.log('Code:');
  console.log('  wrongPara.addText(\'Important Information<w:t xml:space="preserve">1</w:t>\');');
  console.log();
  console.log('What you see in Word:');
  console.log('  "Important Information<w:t xml:space="preserve">1"');
  console.log();
  console.log('Why? The framework correctly escapes XML tags, so they display as literal text.');
  console.log();

  // SECTION 2: The Right Way (Correct Usage)
  console.log('SECTION 2: The RIGHT Way (Correct Usage)');
  console.log('-'.repeat(80));

  const rightDoc = Document.create();
  const rightPara = rightDoc.createParagraph();

  // CORRECT: Use separate text runs or combine text
  rightPara.addText('Important Information');
  rightPara.addText('1');
  // Or simply: rightPara.addText('Important Information 1');

  console.log('Code (Option 1 - Separate Runs):');
  console.log('  rightPara.addText(\'Important Information\');');
  console.log('  rightPara.addText(\'1\');');
  console.log();
  console.log('Code (Option 2 - Combined):');
  console.log('  rightPara.addText(\'Important Information 1\');');
  console.log();
  console.log('What you see in Word:');
  console.log('  "Important Information1" or "Important Information 1"');
  console.log();

  // SECTION 3: Detection Tool
  console.log('SECTION 3: Detecting Corruption with detectCorruptionInDocument()');
  console.log('-'.repeat(80));

  // Detect corruption in the wrong document
  const report = detectCorruptionInDocument(wrongDoc);

  console.log('Corruption Report:');
  console.log(`  Is Corrupted: ${report.isCorrupted}`);
  console.log(`  Total Locations: ${report.totalLocations}`);
  console.log();

  if (report.isCorrupted) {
    console.log('Corruption Details:');
    report.locations.forEach((loc, idx) => {
      console.log(`  Location ${idx + 1}:`);
      console.log(`    Paragraph: ${loc.paragraphIndex}, Run: ${loc.runIndex}`);
      console.log(`    Type: ${loc.corruptionType}`);
      console.log(`    Original Text: "${loc.text.substring(0, 60)}${loc.text.length > 60 ? '...' : ''}"`);
      console.log(`    Suggested Fix: "${loc.suggestedFix}"`);
      console.log();
    });

    console.log('Summary:');
    console.log(report.summary);
    console.log();
  }

  // SECTION 4: Auto-Cleaning Option
  console.log('SECTION 4: Auto-Cleaning with cleanXmlFromText Option');
  console.log('-'.repeat(80));

  const autoCleanDoc = Document.create();
  const autoCleanPara = autoCleanDoc.createParagraph();

  // Enable auto-cleaning to remove XML patterns
  const corruptedText = 'Important Information<w:t xml:space="preserve">1</w:t>';
  autoCleanPara.addText(corruptedText, { cleanXmlFromText: true });

  console.log('Code:');
  console.log('  const corruptedText = \'Important Information<w:t xml:space="preserve">1</w:t>\';');
  console.log('  autoCleanPara.addText(corruptedText, { cleanXmlFromText: true });');
  console.log();
  console.log('What you see in Word:');
  console.log('  "Important Information1" (XML tags removed automatically)');
  console.log();

  // SECTION 5: Common Scenarios
  console.log('SECTION 5: Common Scenarios and Solutions');
  console.log('-'.repeat(80));

  const scenariosDoc = Document.create();

  // Scenario 1: Adding formatted text
  console.log('Scenario 1: Adding formatted text');
  const para1 = scenariosDoc.createParagraph();
  para1.addText('Bold Text', { bold: true });
  para1.addText(' and ');
  para1.addText('Italic Text', { italic: true });
  console.log('  para.addText(\'Bold Text\', { bold: true });');
  console.log('  para.addText(\' and \');');
  console.log('  para.addText(\'Italic Text\', { italic: true });');
  console.log();

  // Scenario 2: Adding hyperlinks
  console.log('Scenario 2: Adding hyperlinks');
  const para2 = scenariosDoc.createParagraph();
  para2.addText('Visit ');
  para2.addHyperlink({
    url: 'https://example.com',
    text: 'our website',
  });
  para2.addText(' for more info.');
  console.log('  para.addText(\'Visit \');');
  console.log('  para.addHyperlink({ url: \'https://example.com\', text: \'our website\' });');
  console.log('  para.addText(\' for more info.\');');
  console.log();

  // Scenario 3: Multiple formatting in one paragraph
  console.log('Scenario 3: Complex formatting');
  const para3 = scenariosDoc.createParagraph();
  para3.addText('Document Title', { bold: true, size: 16 });
  para3.addText(' - Section 1.2.3', { size: 12, color: '808080' });
  console.log('  para.addText(\'Document Title\', { bold: true, size: 16 });');
  console.log('  para.addText(\' - Section 1.2.3\', { size: 12, color: \'808080\' });');
  console.log();

  // Save examples
  const wrongPath = path.join(__dirname, 'output-corrupted.docx');
  const rightPath = path.join(__dirname, 'output-correct.docx');
  const autoCleanPath = path.join(__dirname, 'output-auto-cleaned.docx');
  const scenariosPath = path.join(__dirname, 'output-common-scenarios.docx');

  await wrongDoc.save(wrongPath);
  await rightDoc.save(rightPath);
  await autoCleanDoc.save(autoCleanPath);
  await scenariosDoc.save(scenariosPath);

  console.log('='.repeat(80));
  console.log('Example documents created:');
  console.log(`  ${wrongPath} - Shows corruption (XML tags as text)`);
  console.log(`  ${rightPath} - Correct usage`);
  console.log(`  ${autoCleanPath} - Auto-cleaned text`);
  console.log(`  ${scenariosPath} - Common scenarios with proper formatting`);
  console.log('='.repeat(80));

  // Clean up
  wrongDoc.dispose();
  rightDoc.dispose();
  autoCleanDoc.dispose();
  scenariosDoc.dispose();
}

// Run the demo
demonstrateXmlCorruption().catch(console.error);
