/**
 * Example: Ensure Blank Lines After Tables
 *
 * This example demonstrates how to use the ensureBlankLinesAfter1x1Tables() and
 * ensureBlankLinesAfterOtherTables() methods to add preserved blank lines after
 * tables in a document with configurable styles.
 *
 * This is particularly useful when:
 * 1. You want consistent spacing after Header 2 tables (1x1)
 * 2. You want consistent spacing after multi-cell tables
 * 3. You're processing documents with Template_UI and want to preserve spacing
 * 4. You need to ensure blank lines won't be removed by "remove blank lines" operations
 * 5. You want to use a specific style (e.g., BodyText) instead of Normal
 */

import { Document } from "../src";

async function main() {
  console.log("Loading document...");
  const doc = await Document.load("input.docx");

  console.log("\nBefore processing:");
  const tables = doc.getAllTables();
  console.log(`Total tables: ${tables.length}`);
  const oneByone = tables.filter(
    (t) => t.getRowCount() === 1 && t.getColumnCount() === 1
  );
  console.log(`1x1 tables: ${oneByone.length}`);
  console.log(`Multi-cell tables: ${tables.length - oneByone.length}`);

  // Ensure blank lines after all 1x1 tables
  console.log("\nProcessing 1x1 tables...");
  const result1x1 = doc.ensureBlankLinesAfter1x1Tables({
    spacingAfter: 120, // 6pt spacing (default)
    markAsPreserved: true, // Mark as preserved (default)
    style: "Normal", // Normal style (default)
  });

  console.log("\n1x1 Table Results:");
  console.log(`Tables processed: ${result1x1.tablesProcessed}`);
  console.log(`Blank lines added: ${result1x1.blankLinesAdded}`);
  console.log(`Existing blank lines marked: ${result1x1.existingLinesMarked}`);

  // Ensure blank lines after all multi-cell tables
  console.log("\nProcessing multi-cell tables...");
  const resultOther = doc.ensureBlankLinesAfterOtherTables({
    spacingAfter: 120, // 6pt spacing (default)
    markAsPreserved: true, // Mark as preserved (default)
    style: "Normal", // Normal style (default)
  });

  console.log("\nMulti-cell Table Results:");
  console.log(`Tables processed: ${resultOther.tablesProcessed}`);
  console.log(`Blank lines added: ${resultOther.blankLinesAdded}`);
  console.log(
    `Existing blank lines marked: ${resultOther.existingLinesMarked}`
  );

  // Save the document
  console.log("\nSaving document...");
  await doc.save("output.docx");
  console.log("Done! Document saved to output.docx");

  // Cleanup
  doc.dispose();
}

// Example 2: Only process Header 2 tables with custom style
async function processHeader2TablesOnly() {
  const doc = await Document.load("input.docx");

  const result = doc.ensureBlankLinesAfter1x1Tables({
    style: "BodyText", // Use BodyText style instead of Normal
    filter: (table, index) => {
      const cell = table.getCell(0, 0);
      if (!cell) return false;

      // Check if cell contains a Header 2 paragraph
      return cell.getParagraphs().some((p) => {
        const style = p.getStyle();
        return (
          style === "Heading2" || style === "Heading 2" || style === "Header2"
        );
      });
    },
  });

  console.log(
    `Processed ${result.tablesProcessed} Header 2 tables with BodyText style`
  );
  await doc.save("output.docx");
  doc.dispose();
}

// Example 3: Custom spacing and style for multi-cell tables
async function customStyleForMultiCellTables() {
  const doc = await Document.load("input.docx");

  const result = doc.ensureBlankLinesAfterOtherTables({
    spacingAfter: 240, // 12pt spacing instead of default 6pt
    markAsPreserved: true,
    style: "BodyText", // Use BodyText style instead of Normal
  });

  console.log(
    `Added/marked ${
      result.blankLinesAdded + result.existingLinesMarked
    } blank lines with 12pt spacing and BodyText style`
  );
  await doc.save("output.docx");
  doc.dispose();
}

// Example 4: Different styles for different table types
async function differentStylesForTableTypes() {
  const doc = await Document.load("input.docx");

  // Use 'Normal' for 1x1 tables (Header 2 tables)
  const result1x1 = doc.ensureBlankLinesAfter1x1Tables({
    style: "Normal",
    spacingAfter: 120,
  });

  // Use 'BodyText' for multi-cell tables
  const resultMulti = doc.ensureBlankLinesAfterOtherTables({
    style: "BodyText",
    spacingAfter: 180, // 9pt spacing
  });

  console.log(
    `1x1 tables: Added/marked ${
      result1x1.blankLinesAdded + result1x1.existingLinesMarked
    } blank lines with Normal style`
  );
  console.log(
    `Multi-cell tables: Added/marked ${
      resultMulti.blankLinesAdded + resultMulti.existingLinesMarked
    } blank lines with BodyText style`
  );

  await doc.save("output.docx");
  doc.dispose();
}

// Run main example
main().catch(console.error);
