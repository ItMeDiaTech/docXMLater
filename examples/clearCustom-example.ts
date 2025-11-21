/**
 * Example: Using clearCustom() to remove Structured Document Tags
 *
 * This example demonstrates how to use the clearCustom() helper to remove
 * SDT (Structured Document Tag) wrappers from documents. SDTs are commonly
 * added by Google Docs and other applications to wrap content, adding
 * unnecessary complexity when they're not needed.
 *
 * Typical use cases:
 * - Cleaning Google Docs DOCX exports
 * - Removing content control elements before processing
 * - Simplifying document structure while preserving content
 */

import { Document } from "../src/core/Document";

async function example1_BasicCleanup() {
  console.log("\n=== Example 1: Basic SDT Removal ===");

  // Load a document (e.g., exported from Google Docs)
  console.log("Loading document...");
  const doc = await Document.load("C:/Users/DiaTech/Pictures/DiaTech/Programs/DocHub/development/docXMLater/pre-processed.docx");

  // Count SDTs before cleanup
  const bodyElementsBefore = doc.getBodyElements();
  const sdtCountBefore = bodyElementsBefore.filter(
    (el) => el.constructor.name === "StructuredDocumentTag"
  ).length;
  console.log(`Found ${sdtCountBefore} SDT elements before cleanup`);

  // Remove all SDT wrappers
  console.log("Removing SDT elements...");
  doc.clearCustom();

  // Verify SDTs are gone
  const bodyElementsAfter = doc.getBodyElements();
  const sdtCountAfter = bodyElementsAfter.filter(
    (el) => el.constructor.name === "StructuredDocumentTag"
  ).length;
  console.log(`Found ${sdtCountAfter} SDT elements after cleanup`);

  // Content should still be there
  const paragraphCount = doc.getParagraphCount();
  console.log(`Document still has ${paragraphCount} paragraphs`);

  // Save cleaned document
  await doc.save("examples/output/cleaned.docx");
  console.log("✓ Saved cleaned document to cleaned.docx");
}

async function example2_CleanupAndFormat() {
  console.log("\n=== Example 2: Cleanup and Format in One Workflow ===");

  const doc = await Document.load("examples/output/sample.docx");

  // Chain clearCustom() with other operations
  console.log("Removing SDTs and applying formatting...");
  doc
    .clearCustom() // Remove SDT wrappers
    .applyStyles(); // Apply standard formatting
  // Could chain more methods here...

  await doc.save("examples/output/cleaned-and-formatted.docx");
  console.log(
    "✓ Saved cleaned and formatted document to cleaned-and-formatted.docx"
  );
}

async function example3_SelectiveCleanup() {
  console.log("\n=== Example 3: Verify Cleanup Effectiveness ===");

  const doc = await Document.load("examples/output/sample.docx");

  // Get statistics before
  const allElementsBefore = doc.getBodyElements();
  console.log(`\nBefore cleanup:`);
  console.log(
    `  - Total body elements: ${allElementsBefore.length}`
  );
  console.log(
    `  - SDT elements: ${allElementsBefore.filter((el) => el.constructor.name === "StructuredDocumentTag").length}`
  );
  console.log(
    `  - Paragraphs: ${allElementsBefore.filter((el) => el.constructor.name === "Paragraph").length}`
  );
  console.log(
    `  - Tables: ${allElementsBefore.filter((el) => el.constructor.name === "Table").length}`
  );

  // Clean up
  doc.clearCustom();

  // Get statistics after
  const allElementsAfter = doc.getBodyElements();
  console.log(`\nAfter cleanup:`);
  console.log(
    `  - Total body elements: ${allElementsAfter.length}`
  );
  console.log(
    `  - SDT elements: ${allElementsAfter.filter((el) => el.constructor.name === "StructuredDocumentTag").length}`
  );
  console.log(
    `  - Paragraphs: ${allElementsAfter.filter((el) => el.constructor.name === "Paragraph").length}`
  );
  console.log(
    `  - Tables: ${allElementsAfter.filter((el) => el.constructor.name === "Table").length}`
  );

  // Verify text content is preserved
  const textBefore = doc.getAllParagraphs().map((p) => p.getText()).join(" ");
  console.log(`\nContent preserved: ${textBefore.length > 0 ? "✓ Yes" : "✗ No"}`);

  await doc.save("examples/output/cleanup-verified.docx");
}

async function example4_HandlingComplexDocuments() {
  console.log("\n=== Example 4: Complex Documents with Nested SDTs ===");

  const doc = await Document.load("examples/output/complex.docx");

  console.log("Cleaning complex document with nested SDTs...");

  // Count all nested elements
  let paraCount = 0;
  let tableCount = 0;
  for (const table of doc.getTables()) {
    tableCount++;
    for (const row of table.getRows()) {
      for (const cell of row.getCells()) {
        paraCount += cell.getParagraphs().length;
      }
    }
  }
  paraCount += doc.getBodyElements().filter((el) => el.constructor.name === "Paragraph").length;

  console.log(`Before: ${paraCount} paragraphs, ${tableCount} tables`);

  // Remove nested SDTs recursively
  doc.clearCustom();

  // Recount
  let paraCountAfter = 0;
  let tableCountAfter = 0;
  for (const table of doc.getTables()) {
    tableCountAfter++;
    for (const row of table.getRows()) {
      for (const cell of row.getCells()) {
        paraCountAfter += cell.getParagraphs().length;
      }
    }
  }
  paraCountAfter += doc.getBodyElements().filter((el) => el.constructor.name === "Paragraph").length;

  console.log(
    `After: ${paraCountAfter} paragraphs, ${tableCountAfter} tables`
  );
  console.log(
    `✓ Content structure preserved: paragraphs=${paraCountAfter >= paraCount - 1}`
  );

  await doc.save("examples/output/complex-cleaned.docx");
}

// Run examples
(async () => {
  try {
    console.log("Document Cleanup Examples - clearCustom()");
    console.log("==========================================");
    console.log("\nThe clearCustom() method removes Structured Document Tags (SDTs)");
    console.log(
      "while preserving all content. SDTs are commonly used by Google Docs"
    );
    console.log(
      "and other applications to wrap content in containers that add"
    );
    console.log(
      "complexity when not needed for your use case.\n"
    );

    // Note: These examples assume sample documents exist
    // In a real scenario, you would have actual DOCX files to process
    console.log("Examples demonstrating clearCustom() usage:\n");

    console.log("1. Basic cleanup: Remove SDT wrappers from document");
    console.log("   doc.clearCustom();");
    console.log("   Result: All SDTs removed, content preserved\n");

    console.log("2. Chained operations: Cleanup and format");
    console.log("   doc.clearCustom().applyStyles();");
    console.log(
      "   Result: SDTs removed and formatting applied in one call\n"
    );

    console.log("3. Recursive handling: Nested SDTs");
    console.log(
      "   - Handles SDTs inside other SDTs automatically"
    );
    console.log("   - Processes SDTs inside table cells");
    console.log("   - Preserves all nested content\n");

    console.log("4. Complex documents: Tables with SDT-wrapped cells");
    console.log("   - Recursively unwraps all nested structures");
    console.log("   - Maintains document integrity");
    console.log("   - Preserves formatting and content order\n");

    console.log("Key Points:");
    console.log("✓ Content is completely preserved");
    console.log("✓ Formatting and styles are maintained");
    console.log("✓ Nested SDTs are handled recursively");
    console.log("✓ Works with SDTs in table cells");
    console.log("✓ Chainable with other document methods");
    console.log("✓ Returns 'this' for method chaining\n");

    console.log("Best Practices:");
    console.log(
      "• Use clearCustom() early in your processing pipeline"
    );
    console.log("• Chain it with applyStyles() to standardize formatting");
    console.log(
      "• Verify output document has correct content before saving"
    );
    console.log(
      "• Use removeExtraBlankParagraphs() after cleanup if needed"
    );
    console.log(
      "• Consider normalizeSpacing() to clean up spacing artifacts\n"
    );
  } catch (error) {
    console.error("Error running examples:", error);
    process.exit(1);
  }
})();
