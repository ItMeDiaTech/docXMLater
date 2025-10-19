/**
 * Fix Corrupted Document - Detect and Fix XML Corruption
 *
 * This example shows how to load a corrupted DOCX file, detect XML corruption,
 * and create a fixed version by cleaning the corrupted text.
 */

import { Document, detectCorruptionInDocument, suggestFix, Run } from '../../src/index';
import * as path from 'path';

async function fixCorruptedDocument() {
  console.log('='.repeat(80));
  console.log('Fix Corrupted Document - Detection and Repair');
  console.log('='.repeat(80));
  console.log();

  // STEP 1: Create a corrupted document (for demonstration)
  console.log('STEP 1: Creating a corrupted document for demonstration...');
  console.log('-'.repeat(80));

  const corruptedDoc = Document.create();

  // Simulate the actual corruption from the bug report
  const corruptedTexts = [
    'Important Information&lt;w:t xml:space=&quot;preserve&quot;&gt;1&lt;/w:t&gt;',
    'CVS Specialty Pharmacy Plan Provisions&lt;w:t xml:space=&quot;preserve&quot;&gt;1&lt;/w:t&gt;',
    'CCR Process&lt;w:t xml:space=&quot;preserve&quot;&gt;1&lt;/w:t&gt;',
    'Related Documents&lt;w:t xml:space=&quot;preserve&quot;&gt;1&lt;/w:t&gt;',
  ];

  for (const corruptedText of corruptedTexts) {
    const para = corruptedDoc.createParagraph();
    para.addRun(new Run(corruptedText));
  }

  // Add some clean paragraphs too
  corruptedDoc.createParagraph('This is normal text.');
  corruptedDoc.createParagraph('No corruption here.');

  const corruptedPath = path.join(__dirname, 'corrupted-example.docx');
  await corruptedDoc.save(corruptedPath);
  console.log(`Created corrupted document: ${corruptedPath}`);
  console.log();

  // STEP 2: Load and detect corruption
  console.log('STEP 2: Loading document and detecting corruption...');
  console.log('-'.repeat(80));

  const loadedDoc = await Document.load(corruptedPath);
  const report = detectCorruptionInDocument(loadedDoc);

  console.log(`Is Corrupted: ${report.isCorrupted}`);
  console.log(`Total Corrupted Locations: ${report.totalLocations}`);
  console.log();

  if (report.isCorrupted) {
    console.log('Corruption Statistics:');
    console.log(`  Escaped XML: ${report.statistics.escapedXml}`);
    console.log(`  XML Tags: ${report.statistics.xmlTags}`);
    console.log(`  Entities: ${report.statistics.entities}`);
    console.log(`  Mixed: ${report.statistics.mixed}`);
    console.log();

    // STEP 3: Display corruption details
    console.log('STEP 3: Corruption details:');
    console.log('-'.repeat(80));

    report.locations.forEach((loc, idx) => {
      console.log(`\nLocation ${idx + 1}:`);
      console.log(`  Paragraph: ${loc.paragraphIndex}`);
      console.log(`  Run: ${loc.runIndex}`);
      console.log(`  Type: ${loc.corruptionType}`);
      console.log(`  Length: ${loc.length} characters`);
      console.log();
      console.log(`  Original Text:`);
      console.log(`    "${loc.text}"`);
      console.log();
      console.log(`  Suggested Fix:`);
      console.log(`    "${loc.suggestedFix}"`);
    });
    console.log();

    // STEP 4: Create fixed document
    console.log('STEP 4: Creating fixed document...');
    console.log('-'.repeat(80));

    const fixedDoc = Document.create();
    const paragraphs = loadedDoc.getParagraphs();

    for (let pIdx = 0; pIdx < paragraphs.length; pIdx++) {
      const originalPara = paragraphs[pIdx];
      if (!originalPara) continue;

      const newPara = fixedDoc.createParagraph();

      // Copy paragraph formatting
      const formatting = originalPara.getFormatting();
      if (formatting.alignment) newPara.setAlignment(formatting.alignment);
      if (formatting.style) newPara.setStyle(formatting.style);

      // Copy and fix runs
      const runs = originalPara.getRuns();
      for (let rIdx = 0; rIdx < runs.length; rIdx++) {
        const run = runs[rIdx];
        if (!run) continue;

        const originalText = run.getText();
        const runFormatting = run.getFormatting();

        // Check if this run is corrupted
        const corruptionLoc = report.locations.find(
          loc => loc.paragraphIndex === pIdx && loc.runIndex === rIdx
        );

        // Use fixed text if corrupted, otherwise use original
        const text = corruptionLoc ? corruptionLoc.suggestedFix : originalText;

        // Add the fixed/original run
        newPara.addText(text, runFormatting);
      }
    }

    // Save fixed document
    const fixedPath = path.join(__dirname, 'fixed-example.docx');
    await fixedDoc.save(fixedPath);

    console.log(`Fixed document saved: ${fixedPath}`);
    console.log();
    console.log('Changes made:');
    report.locations.forEach((loc, idx) => {
      console.log(`  ${idx + 1}. Paragraph ${loc.paragraphIndex}, Run ${loc.runIndex}:`);
      console.log(`     Before: "${loc.text.substring(0, 50)}${loc.text.length > 50 ? '...' : ''}"`);
      console.log(`     After:  "${loc.suggestedFix}"`);
    });

    // Clean up
    fixedDoc.dispose();
  } else {
    console.log('No corruption detected in the document.');
  }

  // STEP 5: Show comparison
  console.log();
  console.log('STEP 5: Text comparison:');
  console.log('-'.repeat(80));

  const paragraphs = loadedDoc.getParagraphs();
  paragraphs.forEach((para, idx) => {
    const text = para.getText();
    if (text.length > 0) {
      console.log(`Paragraph ${idx}:`);
      console.log(`  "${text}"`);

      // Check if this paragraph has corruption
      const hasCorruption = report.locations.some(loc => loc.paragraphIndex === idx);
      if (hasCorruption) {
        console.log('  ⚠️  Contains corruption');
      } else {
        console.log('  ✅ Clean');
      }
      console.log();
    }
  });

  console.log('='.repeat(80));
  console.log('Summary:');
  console.log(report.summary);
  console.log('='.repeat(80));

  // Clean up
  corruptedDoc.dispose();
  loadedDoc.dispose();
}

async function demonstrateManualCleaning() {
  console.log();
  console.log('='.repeat(80));
  console.log('Manual Text Cleaning Example');
  console.log('='.repeat(80));
  console.log();

  const examples = [
    'Important Information&lt;w:t xml:space=&quot;preserve&quot;&gt;1&lt;/w:t&gt;',
    'Text with <w:r><w:t>embedded</w:t></w:r> XML',
    'Mixed &lt;w:p&gt; and <w:t> patterns',
  ];

  examples.forEach((example, idx) => {
    const fixed = suggestFix(example);
    console.log(`Example ${idx + 1}:`);
    console.log(`  Original: "${example}"`);
    console.log(`  Fixed:    "${fixed}"`);
    console.log();
  });

  console.log('='.repeat(80));
}

// Run both demonstrations
(async () => {
  await fixCorruptedDocument();
  await demonstrateManualCleaning();
})().catch(console.error);
