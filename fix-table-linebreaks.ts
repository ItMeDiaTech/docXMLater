/**
 * Fix Table Linebreaks
 *
 * This script ensures all 1x1 tables have a linebreak (blank paragraph) after them.
 * This fixes documents where tables are missing proper spacing.
 *
 * Usage:
 *   npx ts-node fix-table-linebreaks.ts input.docx output.docx
 */

import { Document } from './src/core/Document';
import { Table } from './src/elements/Table';
import { Paragraph } from './src/elements/Paragraph';
import * as fs from 'fs';
import * as path from 'path';

interface TableFix {
  tableIndex: number;
  hadLinebreak: boolean;
  addedLinebreak: boolean;
  cellContent: string;
}

async function fixTableLinebreaks(inputPath: string, outputPath: string): Promise<void> {
  console.log(`\n🔧 Fixing Table Linebreaks`);
  console.log('═'.repeat(60));
  console.log(`Input:  ${path.basename(inputPath)}`);
  console.log(`Output: ${path.basename(outputPath)}\n`);

  // Load document
  const doc = await Document.load(inputPath);
  const bodyElements = (doc as any).bodyElements || [];

  const fixes: TableFix[] = [];
  let addedCount = 0;
  let skippedCount = 0;

  // Iterate through body elements
  for (let i = 0; i < bodyElements.length; i++) {
    const element = bodyElements[i];

    // Check if it's a table
    if (!(element instanceof Table)) {
      continue;
    }

    const table = element as Table;
    const rows = table.getRowCount();
    const cols = table.getColumnCount();

    // Only process 1x1 tables
    if (rows !== 1 || cols !== 1) {
      continue;
    }

    // Get cell content for logging
    let cellContent = '';
    const cell = table.getCell(0, 0);
    if (cell) {
      const paras = cell.getParagraphs();
      cellContent = paras.map(p => p.getText()).join(' ').trim();
    }

    // Check if next element exists and is a blank paragraph
    const nextElement = bodyElements[i + 1];
    let hasLinebreak = false;

    if (nextElement instanceof Paragraph) {
      const text = nextElement.getText().trim();
      if (text === '') {
        hasLinebreak = true;
      }
    }

    // Track fix
    const fix: TableFix = {
      tableIndex: i,
      hadLinebreak: hasLinebreak,
      addedLinebreak: false,
      cellContent: cellContent.substring(0, 50) + (cellContent.length > 50 ? '...' : '')
    };

    // Add linebreak if missing
    if (!hasLinebreak) {
      const blankPara = Paragraph.create();

      // Add spacing to ensure visibility in Word
      blankPara.setSpaceAfter(120); // 120 twips = 6pt

      // Mark as preserved so it won't be removed by cleanup operations
      blankPara.setPreserved(true);

      // Insert after table
      bodyElements.splice(i + 1, 0, blankPara);

      fix.addedLinebreak = true;
      addedCount++;

      console.log(`✓ Added linebreak after table ${i}`);
      console.log(`  Content: "${fix.cellContent}"`);

      // Skip the newly added paragraph in the next iteration
      i++;
    } else {
      skippedCount++;
      console.log(`⊘ Skipped table ${i} (already has linebreak)`);
      console.log(`  Content: "${fix.cellContent}"`);
    }

    fixes.push(fix);
  }

  // Save the fixed document
  await doc.save(outputPath);

  // Summary
  console.log('\n' + '═'.repeat(60));
  console.log(`📊 SUMMARY:`);
  console.log(`  Total 1×1 tables found: ${fixes.length}`);
  console.log(`  Linebreaks added: ${addedCount}`);
  console.log(`  Already had linebreaks: ${skippedCount}`);
  console.log(`\n✅ Fixed document saved to: ${outputPath}\n`);
}

// Main execution
if (require.main === module) {
  const args = process.argv.slice(2);

  if (args.length !== 2) {
    console.log(`
Usage: npx ts-node fix-table-linebreaks.ts <input.docx> <output.docx>

Example:
  npx ts-node fix-table-linebreaks.ts BEFORE.docx AFTER.docx

This script:
  1. Finds all 1×1 tables in the document
  2. Checks if each table has a blank paragraph (linebreak) after it
  3. Adds a blank paragraph if missing
  4. Saves the fixed document
    `);
    process.exit(1);
  }

  const [inputPath, outputPath] = args;

  // Check input exists
  if (!fs.existsSync(inputPath)) {
    console.error(`❌ Input file not found: ${inputPath}`);
    process.exit(1);
  }

  fixTableLinebreaks(inputPath, outputPath)
    .then(() => {
      console.log('✅ Done!\n');
    })
    .catch((error) => {
      console.error('\n❌ Error:', error);
      process.exit(1);
    });
}

export { fixTableLinebreaks };
