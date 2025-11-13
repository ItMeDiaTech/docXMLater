/**
 * Table Linebreak Analysis Tool
 *
 * This script analyzes DOCX documents to identify why linebreaks aren't
 * appearing after 1x1 tables. It compares BEFORE and AFTER documents
 * and shows the structure differences.
 *
 * Usage:
 *   npx ts-node analyze-table-linebreaks.ts BEFORE.docx AFTER.docx
 */

import { Document } from './src/core/Document';
import { Table } from './src/elements/Table';
import { Paragraph } from './src/elements/Paragraph';
import * as fs from 'fs';
import * as path from 'path';

interface TableAnalysis {
  index: number;
  is1x1: boolean;
  rows: number;
  cols: number;
  cellContent: string;
  hasLinebreakAfter: boolean;
  nextElementType: string;
  nextElementContent: string;
  tableFormatting: any;
}

async function analyzeDocument(filePath: string): Promise<{
  tables: TableAnalysis[];
  summary: string;
}> {
  console.log(`\n📄 Analyzing: ${path.basename(filePath)}`);
  console.log('═'.repeat(60));

  const doc = await Document.load(filePath);
  const allElements = (doc as any).bodyElements || [];
  const tables: TableAnalysis[] = [];

  console.log(`Total body elements: ${allElements.length}`);

  // Analyze each table
  allElements.forEach((element: any, index: number) => {
    if (element instanceof Table) {
      const rows = element.getRowCount();
      const cols = element.getColumnCount();
      const is1x1 = rows === 1 && cols === 1;

      // Get cell content
      let cellContent = '';
      if (is1x1) {
        const cell = element.getCell(0, 0);
        if (cell) {
          const paras = cell.getParagraphs();
          cellContent = paras.map((p: Paragraph) => p.getText()).join(' ').trim();
        }
      }

      // Check next element
      const nextElement = allElements[index + 1];
      let hasLinebreakAfter = false;
      let nextElementType = 'none';
      let nextElementContent = '';

      if (nextElement) {
        nextElementType = nextElement.constructor.name;
        if (nextElement instanceof Paragraph) {
          nextElementContent = nextElement.getText().trim();
          // A linebreak would be an empty paragraph
          if (nextElementContent === '') {
            hasLinebreakAfter = true;
          }
        }
      }

      const analysis: TableAnalysis = {
        index,
        is1x1,
        rows,
        cols,
        cellContent,
        hasLinebreakAfter,
        nextElementType,
        nextElementContent,
        tableFormatting: element.getFormatting(),
      };

      tables.push(analysis);

      // Print analysis
      console.log(`\n📊 Table ${tables.length} (Element ${index}):`);
      console.log(`   Dimensions: ${rows}×${cols} ${is1x1 ? '✓ (1×1)' : ''}`);
      if (is1x1) {
        console.log(`   Content: "${cellContent.substring(0, 50)}${cellContent.length > 50 ? '...' : ''}"`);
      }
      console.log(`   Next element: ${nextElementType}`);
      if (nextElementType === 'Paragraph') {
        console.log(`   Next content: "${nextElementContent.substring(0, 50)}${nextElementContent.length > 50 ? '...' : ''}"`);
      }
      console.log(`   Has linebreak after: ${hasLinebreakAfter ? '✓ YES' : '✗ NO'}`);
    }
  });

  // Summary
  const total1x1Tables = tables.filter(t => t.is1x1).length;
  const tablesWithLinebreak = tables.filter(t => t.is1x1 && t.hasLinebreakAfter).length;
  const tablesWithoutLinebreak = tables.filter(t => t.is1x1 && !t.hasLinebreakAfter).length;

  const summary = `
Summary for ${path.basename(filePath)}:
  • Total tables: ${tables.length}
  • 1×1 tables: ${total1x1Tables}
  • 1×1 tables WITH linebreak after: ${tablesWithLinebreak}
  • 1×1 tables WITHOUT linebreak after: ${tablesWithoutLinebreak}
`;

  console.log('\n' + summary);

  return { tables, summary };
}

async function compareDocuments(beforePath: string, afterPath: string) {
  console.log('🔍 Table Linebreak Analysis Tool');
  console.log('═'.repeat(60));

  // Check files exist
  if (!fs.existsSync(beforePath)) {
    console.error(`❌ File not found: ${beforePath}`);
    process.exit(1);
  }
  if (!fs.existsSync(afterPath)) {
    console.error(`❌ File not found: ${afterPath}`);
    process.exit(1);
  }

  // Analyze both documents
  const beforeAnalysis = await analyzeDocument(beforePath);
  const afterAnalysis = await analyzeDocument(afterPath);

  // Compare results
  console.log('\n\n🔄 COMPARISON');
  console.log('═'.repeat(60));

  const before1x1 = beforeAnalysis.tables.filter(t => t.is1x1);
  const after1x1 = afterAnalysis.tables.filter(t => t.is1x1);

  console.log(`\nBEFORE: ${before1x1.length} 1×1 tables`);
  console.log(`AFTER:  ${after1x1.length} 1×1 tables`);

  // Check if linebreaks were added/removed
  const beforeWithoutLinebreak = before1x1.filter(t => !t.hasLinebreakAfter).length;
  const afterWithoutLinebreak = after1x1.filter(t => !t.hasLinebreakAfter).length;

  console.log(`\nTables missing linebreaks:`);
  console.log(`  BEFORE: ${beforeWithoutLinebreak}`);
  console.log(`  AFTER:  ${afterWithoutLinebreak}`);

  if (afterWithoutLinebreak > 0) {
    console.log(`\n⚠️  ISSUE: ${afterWithoutLinebreak} table(s) still missing linebreaks after processing`);
  } else if (beforeWithoutLinebreak > 0 && afterWithoutLinebreak === 0) {
    console.log(`\n✅ SUCCESS: All linebreaks added!`);
  } else {
    console.log(`\n✅ All 1×1 tables have linebreaks`);
  }

  // Detailed comparison for problematic tables
  const problematicAfter = after1x1.filter(t => !t.hasLinebreakAfter);
  if (problematicAfter.length > 0) {
    console.log(`\n\n🔍 PROBLEMATIC TABLES (Missing Linebreaks):`);
    console.log('═'.repeat(60));

    problematicAfter.forEach((table, idx) => {
      console.log(`\n${idx + 1}. Element ${table.index}:`);
      console.log(`   Content: "${table.cellContent.substring(0, 60)}${table.cellContent.length > 60 ? '...' : ''}"`);
      console.log(`   Next element: ${table.nextElementType}`);
      if (table.nextElementType === 'Paragraph') {
        console.log(`   Next content: "${table.nextElementContent.substring(0, 60)}${table.nextElementContent.length > 60 ? '...' : ''}"`);
      }

      // Check if there's a style issue
      console.log(`\n   💡 Analysis:`);
      if (table.nextElementType === 'Paragraph' && table.nextElementContent !== '') {
        console.log(`      • Next paragraph has content - linebreak should be BEFORE it`);
      } else if (table.nextElementType === 'Table') {
        console.log(`      • Next element is another table - linebreak needed between them`);
      } else if (table.nextElementType === 'none') {
        console.log(`      • Table is at end of document - linebreak should be added`);
      } else {
        console.log(`      • Unexpected situation - needs investigation`);
      }
    });
  }

  // Extract XML for further analysis
  console.log(`\n\n📋 RECOMMENDATIONS:`);
  console.log('═'.repeat(60));
  console.log(`
1. Check if blank paragraphs are being removed during processing
2. Verify that wrapParagraphInTable() adds blank paragraphs correctly
3. Ensure removeExtraBlankParagraphs() preserves linebreaks after tables
4. Check if Header2 table processing is working as expected
  `);

  // Generate output report
  const reportPath = 'table-linebreak-analysis-report.txt';
  const report = `
TABLE LINEBREAK ANALYSIS REPORT
Generated: ${new Date().toISOString()}

${beforeAnalysis.summary}
${afterAnalysis.summary}

DETAILED FINDINGS:
${problematicAfter.length > 0 ?
  problematicAfter.map((t, i) => `
${i + 1}. Element ${t.index}:
   Content: "${t.cellContent}"
   Next element: ${t.nextElementType}
   Next content: "${t.nextElementContent}"
`).join('\n') :
  'No issues found - all tables have proper linebreaks'}

RECOMMENDATIONS:
1. Review Document.ts line 3111 - condition may be inverted
2. Check if blank paragraphs are marked as 'preserved' correctly
3. Verify removeExtraBlankParagraphs() isn't removing table linebreaks
4. Test with wrapParagraphInTable() to ensure blank paragraph is added
`;

  fs.writeFileSync(reportPath, report);
  console.log(`\n📄 Full report saved to: ${reportPath}`);
}

// Main execution
if (require.main === module) {
  const args = process.argv.slice(2);

  if (args.length !== 2) {
    console.log(`
Usage: npx ts-node analyze-table-linebreaks.ts <BEFORE.docx> <AFTER.docx>

Example:
  npx ts-node analyze-table-linebreaks.ts BEFORE.docx AFTER.docx
    `);
    process.exit(1);
  }

  const [beforePath, afterPath] = args;

  compareDocuments(beforePath, afterPath)
    .then(() => {
      console.log('\n✅ Analysis complete!\n');
    })
    .catch((error) => {
      console.error('\n❌ Error during analysis:', error);
      process.exit(1);
    });
}
