/**
 * Example: Remove All Preserve Flags
 *
 * Demonstrates how to use the removeAllPreserveFlags() helper method
 * to clear preserve flags from all paragraphs in the document.
 *
 * Preserve flags are runtime-only markers that prevent paragraphs from being
 * automatically removed by cleanup operations. This example shows how to:
 * 1. Mark paragraphs as preserved during document processing
 * 2. Remove all preserve flags when ready for cleanup
 * 3. Remove extra blank paragraphs after clearing preserve flags
 */

import { Document } from "../src/index";

async function removeAllPreserveFlagsExample() {
  // Create a new document
  const doc = Document.create();

  // Add some content with blank lines
  doc.createParagraph("Chapter 1");
  doc.createParagraph(""); // Blank line
  doc.createParagraph(""); // Blank line
  doc.createParagraph("Section 1.1");
  doc.createParagraph(""); // Blank line
  doc.createParagraph("Content here");
  doc.createParagraph(""); // Blank line
  doc.createParagraph(""); // Blank line
  doc.createParagraph(""); // Blank line

  console.log(`Before: ${doc.getAllParagraphs().length} paragraphs`);

  // Mark some blank lines as preserved (to protect them during cleanup)
  const paragraphs = doc.getAllParagraphs();
  const blankParagraphs = paragraphs.filter((p) => p.getText().trim() === "");
  for (const blankPara of blankParagraphs) {
    blankPara.setPreserved(true);
  }

  console.log(`Preserved: ${blankParagraphs.length} blank paragraphs`);

  // At this point, blank paragraphs won't be removed because they're marked as preserved
  const result1 = doc.removeExtraBlankParagraphs();
  console.log(
    `After removeExtraBlankParagraphs (with preserved): Removed ${result1.removed}, Preserved ${result1.preserved}`
  );

  // Now remove all preserve flags
  const cleared = doc.removeAllPreserveFlags();
  console.log(`Cleared preserve flags from ${cleared} paragraphs`);

  // Now blank paragraphs can be removed
  const result2 = doc.removeExtraBlankParagraphs();
  console.log(
    `After removeExtraBlankParagraphs (preserve flags cleared): Removed ${result2.removed}, Preserved ${result2.preserved}`
  );

  console.log(`After: ${doc.getAllParagraphs().length} paragraphs`);

  // Save the document
  await doc.save("removeAllPreserveFlags-example.docx");
  console.log("✅ Document saved as removeAllPreserveFlags-example.docx");
}

// Run the example
removeAllPreserveFlagsExample().catch(console.error);
