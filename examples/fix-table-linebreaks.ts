/**
 * Example: Fixing Missing Linebreaks After 1x1 Tables
 *
 * This example demonstrates how to use the ensureTableLinebreaks() method
 * to automatically add blank paragraphs (linebreaks) after all 1×1 tables
 * in a document.
 *
 * Common Issue:
 * - 1×1 tables are often used for formatting headers or highlighted text
 * - Without linebreaks after them, content appears cramped in Word
 * - Manual fixing is tedious for documents with many tables
 *
 * Solution:
 * - Use doc.ensureTableLinebreaks() to automatically fix all tables
 * - The method adds blank paragraphs only where needed
 * - Existing linebreaks are preserved (not duplicated)
 */

import { Document } from '../src/core/Document';
import * as path from 'path';

async function main() {
  console.log('\n📋 Example: Fixing Table Linebreaks\n');

  // Example 1: Basic Usage
  console.log('1️⃣  Basic Usage - Fix all 1×1 tables:');
  {
    const doc = await Document.load('input.docx');

    // Scan document and add linebreaks after all 1×1 tables
    const result = doc.ensureTableLinebreaks();

    console.log(`   ✓ Total 1×1 tables: ${result.total}`);
    console.log(`   ✓ Linebreaks added: ${result.added}`);
    console.log(`   ✓ Already had linebreaks: ${result.skipped}`);

    await doc.save('output.docx');
    console.log('   ✓ Saved to output.docx\n');
  }

  // Example 2: Custom Spacing
  console.log('2️⃣  Custom Spacing - Add more space after tables:');
  {
    const doc = await Document.load('input.docx');

    // Add linebreaks with 12pt spacing (instead of default 6pt)
    const result = doc.ensureTableLinebreaks({
      spacingAfter: 240, // 240 twips = 12pt
    });

    console.log(`   ✓ Added ${result.added} linebreaks with 12pt spacing\n`);

    await doc.save('output-custom-spacing.docx');
  }

  // Example 3: Selective Processing with Filter
  console.log('3️⃣  Selective Processing - Only fix specific tables:');
  {
    const doc = await Document.load('input.docx');

    // Only add linebreaks after tables containing "Header" text
    const result = doc.ensureTableLinebreaks({
      filter: (table, index) => {
        const cell = table.getCell(0, 0);
        if (!cell) return false;

        const text = cell.getParagraphs()[0]?.getText() || '';
        return text.includes('Header');
      },
    });

    console.log(`   ✓ Processed ${result.total} tables with "Header" text`);
    console.log(`   ✓ Added ${result.added} linebreaks\n`);

    await doc.save('output-selective.docx');
  }

  // Example 4: Don't Mark as Preserved
  console.log('4️⃣  Unpreserved Linebreaks - Allow cleanup to remove them:');
  {
    const doc = await Document.load('input.docx');

    // Add linebreaks but don't mark them as "preserved"
    // This allows removeExtraBlankParagraphs() to remove them if needed
    doc.ensureTableLinebreaks({
      markAsPreserved: false,
    });

    console.log('   ✓ Added linebreaks (not marked as preserved)\n');

    await doc.save('output-unpreserved.docx');
  }

  // Example 5: Integration with Document Processing
  console.log('5️⃣  Full Document Processing Pipeline:');
  {
    const doc = await Document.load('input.docx');

    // Step 1: Apply custom styles
    console.log('   → Applying custom styles...');
    doc.applyDocHubStylesToDocument();

    // Step 2: Ensure table linebreaks
    console.log('   → Ensuring table linebreaks...');
    const linebreaksResult = doc.ensureTableLinebreaks();
    console.log(`      Added ${linebreaksResult.added} linebreaks`);

    // Step 3: Clean up extra blank paragraphs (but preserve table linebreaks)
    console.log('   → Cleaning up extra blank paragraphs...');
    const cleanupResult = doc.removeExtraBlankParagraphs({
      keepOne: true,
      preserveHeader2BlankLines: true,
    });
    console.log(`      Removed ${cleanupResult.removed} extra blanks`);
    console.log(`      Preserved ${cleanupResult.preserved} important blanks`);

    await doc.save('output-processed.docx');
    console.log('   ✓ Complete!\n');
  }

  console.log('✅ All examples completed!\n');
  console.log('📝 Key Takeaways:');
  console.log('   • ensureTableLinebreaks() automatically fixes spacing after 1×1 tables');
  console.log('   • Existing linebreaks are preserved (not duplicated)');
  console.log('   • Linebreaks are marked as "preserved" by default');
  console.log('   • Use with removeExtraBlankParagraphs() for best results');
  console.log('   • Filter option allows selective processing\n');
}

// Run if executed directly
if (require.main === module) {
  main().catch((error) => {
    console.error('❌ Error:', error.message);
    process.exit(1);
  });
}

export { main };
