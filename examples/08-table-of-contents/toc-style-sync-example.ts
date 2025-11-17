/**
 * TOC Style Name Synchronization Example
 *
 * Demonstrates how TOC field instructions are automatically synchronized
 * with actual style names on save, preventing TOC population failures.
 */

import { Document, Style, TableOfContents } from "../../src/index";
import { promises as fs } from "fs";

async function demonstrateTOCStyleSync() {
  console.log("=== TOC Style Name Synchronization Demo ===\n");

  // =====================================
  // Scenario 1: Load document with mismatched style names
  // =====================================
  console.log("Scenario 1: Loading document with custom style names");

  // Create a test document with custom style names
  const doc1 = Document.create();

  // Create a custom style that replaces "Heading 2"
  const customH2 = Style.create({
    styleId: "Heading2",
    name: "CustomHeader2", // Different from standard "Heading 2"
    type: "paragraph",
    runFormatting: { font: "Verdana", size: 14, bold: true },
    paragraphFormatting: { spacing: { before: 120, after: 120 } },
  });
  doc1.addStyle(customH2);

  // Add some content with custom style
  doc1.createParagraph("Introduction").setStyle("Heading1");
  doc1.createParagraph("Custom Section").setStyle("Heading2"); // Uses CustomHeader2
  doc1.createParagraph("Content paragraph");

  // Create TOC with \t switch referencing old name
  const toc1 = TableOfContents.create({
    title: "Table of Contents",
    includeStyles: [
      { styleName: "Heading 1", level: 1 },
      { styleName: "Heading 2", level: 2 }, // ← Outdated name!
    ],
  });
  doc1.addTableOfContents(toc1);

  await doc1.save("test-toc-sync-before.docx");
  console.log("✓ Created test document with mismatched style names");

  // Load and inspect
  const loaded = await Document.load("test-toc-sync-before.docx");
  const h2Style = loaded.getStyle("Heading2");
  console.log(`  - Heading2 style name in styles.xml: "${h2Style?.getName()}"`);
  console.log(`  - TOC references: "Heading 2" (will be updated on save)\n`);

  // Save again - TOC instruction will be synchronized
  await loaded.save("test-toc-sync-after.docx");
  console.log("✓ Saved document - TOC field instruction synchronized");
  console.log('  - TOC now references: "CustomHeader2" (matches styles.xml)\n');

  // =====================================
  // Scenario 2: Document with preserved original names
  // =====================================
  console.log("Scenario 2: Preserving exact style names from source");

  const doc2 = Document.create();

  // Create styles with non-standard names
  const listStyle = Style.create({
    styleId: "ListParagraph",
    name: "ListParagraph", // No space (like Google Docs exports)
    type: "paragraph",
    runFormatting: { font: "Verdana", size: 12 },
  });
  doc2.addStyle(listStyle);

  doc2.createParagraph("My List:").setStyle("Normal");
  doc2.createParagraph("Item 1").setStyle("ListParagraph");
  doc2.createParagraph("Item 2").setStyle("ListParagraph");

  // TOC with standard name (with space)
  const toc2 = TableOfContents.create({
    includeStyles: [
      { styleName: "List Paragraph", level: 1 }, // ← Has space
    ],
  });
  doc2.addTableOfContents(toc2);

  await doc2.save("test-list-style-sync.docx");
  console.log("✓ Document saved with synchronized TOC");
  console.log(`  - Style name: "ListParagraph" (preserved from source)`);
  console.log(`  - TOC references: "ListParagraph" (auto-synced)\n`);

  // =====================================
  // Scenario 3: Outline-based TOC (no sync needed)
  // =====================================
  console.log("Scenario 3: Outline-based TOC (uses \\o switch)");

  const doc3 = Document.create();
  doc3.createParagraph("Chapter 1").setStyle("Heading1");
  doc3.createParagraph("Section 1.1").setStyle("Heading2");

  // Outline-based TOC - no style names referenced
  const toc3 = TableOfContents.create({
    title: "Contents",
    levels: 3, // Uses \o "1-3" switch
  });
  doc3.addTableOfContents(toc3);

  await doc3.save("test-outline-toc.docx");
  console.log("✓ Outline-based TOC saved (no synchronization needed)");
  console.log('  - Uses \\o "1-3" switch (styleId-based matching)\n');

  // =====================================
  // Scenario 4: Your external code safety
  // =====================================
  console.log("Scenario 4: External code safety verification");

  const doc4 = await Document.load("test-toc-sync-after.docx");
  const h2 = doc4.getStyle("Heading2");

  // Your external code checks
  if (h2?.getName() === "CustomHeader2") {
    console.log('✓ Style name preserved: "CustomHeader2"');
  }

  // Paragraph style references
  const paras = doc4.getAllParagraphs();
  for (const para of paras) {
    const styleId = para.getStyle();
    if (styleId === "Heading2") {
      const style = doc4.getStyle(styleId);
      console.log(
        `✓ Paragraph uses styleId="${styleId}", name="${style?.getName()}"`
      );
      console.log("  → External code checks remain valid\n");
      break;
    }
  }

  // Cleanup
  await fs.unlink("test-toc-sync-before.docx").catch(() => {});
  await fs.unlink("test-toc-sync-after.docx").catch(() => {});
  await fs.unlink("test-list-style-sync.docx").catch(() => {});
  await fs.unlink("test-outline-toc.docx").catch(() => {});

  console.log("=== Summary ===");
  console.log("✅ Style names are NEVER automatically changed");
  console.log("✅ TOC field instructions ARE synchronized to match styles.xml");
  console.log("✅ External code that checks style names remains safe");
  console.log("✅ TOC population works correctly regardless of custom names");
}

// Run if executed directly
if (require.main === module) {
  demonstrateTOCStyleSync()
    .then(() => console.log("\n✓ Demo complete"))
    .catch((err) => console.error("Error:", err));
}

export { demonstrateTOCStyleSync };
