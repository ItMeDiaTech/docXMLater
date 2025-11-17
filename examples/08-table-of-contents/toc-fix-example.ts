/**
 * Table of Contents - Fix Example
 * Demonstrates how to properly create and manage TOC in documents
 * with proper error handling and logging
 */

import {
  Document,
  Paragraph,
  TableOfContents,
  ConsoleLogger,
  SilentLogger,
  LogLevel,
} from "../../src/index";
import * as path from "path";

const OUTPUT_DIR = path.join(__dirname, "output");

/**
 * Example 1: Basic TOC with Error Handling
 * Shows proper way to create a TOC with validation
 */
async function example1_BasicTOCWithLogging() {
  console.log("\n=== Example 1: Basic TOC with Logging ===");

  // Use verbose logging to see what's happening
  const logger = new ConsoleLogger(LogLevel.INFO);
  const doc = Document.create({ logger });

  // Add title
  doc.createParagraph("Document with Table of Contents").setStyle("Title");

  // Add TOC
  const toc = TableOfContents.create({
    title: "Table of Contents",
    showPageNumbers: true,
  });
  doc.addTableOfContents(toc);

  // Add some headings
  doc.createParagraph("Introduction").setStyle("Heading1");
  doc.createParagraph("This is the introduction section with some content.");

  doc.createParagraph("Background").setStyle("Heading1");
  doc.createParagraph("This section provides background information.");

  doc.createParagraph("Technical Details").setStyle("Heading2");
  doc.createParagraph("Here are the technical specifications.");

  doc.createParagraph("Conclusion").setStyle("Heading1");
  doc.createParagraph("Final thoughts and summary.");

  await doc.save(path.join(OUTPUT_DIR, "toc-basic-with-logging.docx"));
  console.log("✓ Created document with TOC and logging");
}

/**
 * Example 2: TOC Without Console Noise
 * Use SilentLogger when you don't want framework output
 */
async function example2_TOCSilentMode() {
  console.log("\n=== Example 2: TOC with Silent Logging ===");

  // Suppress all framework logging
  const doc = Document.create({
    logger: new SilentLogger(),
  });

  doc.createParagraph("Silent Mode Document").setStyle("Title");

  const toc = TableOfContents.create();
  doc.addTableOfContents(toc);

  // Add multiple sections
  for (let i = 1; i <= 5; i++) {
    doc.createParagraph(`Section ${i}`).setStyle("Heading1");
    doc.createParagraph(`Content for section ${i}`);

    // Add subsections
    for (let j = 1; j <= 3; j++) {
      doc.createParagraph(`Subsection ${i}.${j}`).setStyle("Heading2");
      doc.createParagraph(`Details for subsection ${i}.${j}`);
    }
  }

  await doc.save(path.join(OUTPUT_DIR, "toc-silent-mode.docx"));
  console.log("✓ Created document with TOC (silent mode)");
}

/**
 * Example 3: TOC with Custom Styling
 * Shows how to customize TOC appearance
 */
async function example3_TOCWithCustomStyling() {
  console.log("\n=== Example 3: TOC with Custom Styling ===");

  const doc = Document.create();

  // Custom title
  doc
    .createParagraph("Advanced Document")
    .setStyle("Title")
    .setAlignment("center");

  // Custom TOC
  const toc = TableOfContents.create({
    title: "Contents",
    showPageNumbers: true,
    rightAlignPageNumbers: true,
    useHyperlinks: true,
    levels: 3, // Include heading levels 1-3
  });
  doc.addTableOfContents(toc);

  // Add page break after TOC
  const breakPara = new Paragraph();
  breakPara.setPageBreakBefore(true);
  doc.addParagraph(breakPara);

  // Add structured content
  doc.createParagraph("Chapter 1: Getting Started").setStyle("Heading1");
  doc.createParagraph("Introduction to the system.");

  doc.createParagraph("Installation").setStyle("Heading2");
  doc.createParagraph("How to install the software.");

  doc.createParagraph("Configuration").setStyle("Heading2");
  doc.createParagraph("Setting up your environment.");

  doc.createParagraph("Chapter 2: Advanced Usage").setStyle("Heading1");
  doc.createParagraph("Deep dive into features.");

  doc.createParagraph("Performance Tuning").setStyle("Heading2");
  doc.createParagraph("Optimization techniques.");

  doc.createParagraph("Best Practices").setStyle("Heading2");
  doc.createParagraph("Recommended approaches.");

  await doc.save(path.join(OUTPUT_DIR, "toc-custom-styling.docx"));
  console.log("✓ Created document with custom TOC styling");
}

/**
 * Example 4: Multi-Level TOC
 * Demonstrates hierarchical document structure
 */
async function example4_MultiLevelTOC() {
  console.log("\n=== Example 4: Multi-Level TOC ===");

  const doc = Document.create();

  doc.createParagraph("Technical Documentation").setStyle("Title");

  const toc = TableOfContents.create({
    title: "Table of Contents",
    showPageNumbers: true,
    useHyperlinks: true,
    levels: 4, // Include heading levels 1-4
  });
  doc.addTableOfContents(toc);

  // Level 1
  doc.createParagraph("1. Architecture Overview").setStyle("Heading1");
  doc.createParagraph("High-level system architecture.");

  // Level 2
  doc.createParagraph("1.1 Frontend Components").setStyle("Heading2");
  doc.createParagraph("User interface components.");

  // Level 3
  doc.createParagraph("1.1.1 React Components").setStyle("Heading3");
  doc.createParagraph("React-based UI elements.");

  // Level 4
  doc.createParagraph("1.1.1.1 Button Component").setStyle("Heading4");
  doc.createParagraph("Button implementation details.");

  // More structure
  doc.createParagraph("1.1.2 Vue Components").setStyle("Heading3");
  doc.createParagraph("Vue-based UI elements.");

  doc.createParagraph("1.2 Backend Services").setStyle("Heading2");
  doc.createParagraph("Server-side services.");

  doc.createParagraph("1.2.1 API Layer").setStyle("Heading3");
  doc.createParagraph("REST API implementation.");

  doc.createParagraph("2. Database Design").setStyle("Heading1");
  doc.createParagraph("Database schema and structure.");

  doc.createParagraph("2.1 Tables").setStyle("Heading2");
  doc.createParagraph("Table definitions.");

  doc.createParagraph("2.2 Relationships").setStyle("Heading2");
  doc.createParagraph("Foreign key relationships.");

  await doc.save(path.join(OUTPUT_DIR, "toc-multi-level.docx"));
  console.log("✓ Created document with multi-level TOC");
}

/**
 * Example 5: TOC Update Workflow
 * Shows proper workflow for documents that need TOC updates
 */
async function example5_TOCUpdateWorkflow() {
  console.log("\n=== Example 5: TOC Update Workflow ===");

  const doc = Document.create();

  doc.createParagraph("Living Document").setStyle("Title");
  doc.createParagraph("This document will be updated over time.");

  // Add TOC
  const toc = TableOfContents.create({
    title: "Contents (Auto-Updated)",
    showPageNumbers: true,
  });
  doc.addTableOfContents(toc);

  // Initial content
  doc.createParagraph("Version 1.0 Content").setStyle("Heading1");
  doc.createParagraph("Initial release content.");

  doc.createParagraph("Features").setStyle("Heading2");
  doc.createParagraph("Basic features included.");

  // Save initial version
  await doc.save(path.join(OUTPUT_DIR, "toc-update-workflow.docx"));
  console.log("✓ Created initial document");

  // Note: To update TOC in Word:
  // 1. Open the document in Microsoft Word
  // 2. Click anywhere in the TOC
  // 3. Press F9 or right-click and select "Update Field"
  // 4. Choose "Update entire table"

  console.log("\nNote: TOC will update automatically when opened in Word");
  console.log(
    'Or manually: Click TOC → Press F9 → Select "Update entire table"'
  );
}

/**
 * Example 6: TOC Best Practices
 * Demonstrates recommended patterns
 */
async function example6_TOCBestPractices() {
  console.log("\n=== Example 6: TOC Best Practices ===");

  const doc = Document.create();

  // 1. Clear title
  doc
    .createParagraph("Professional Report")
    .setStyle("Title")
    .setAlignment("center");

  // 2. Add blank line after title
  doc.createParagraph("");

  // 3. Configure TOC properly
  const toc = TableOfContents.create({
    title: "Table of Contents",
    showPageNumbers: true,
    rightAlignPageNumbers: true,
    useHyperlinks: true,
    levels: 3, // Don't go too deep - include heading levels 1-3
  });
  doc.addTableOfContents(toc);

  // 4. Page break after TOC
  const pageBreak = new Paragraph();
  pageBreak.setPageBreakBefore(true);
  doc.addParagraph(pageBreak);

  // 5. Use consistent heading hierarchy
  doc.createParagraph("Executive Summary").setStyle("Heading1");
  doc.createParagraph("Summary of findings and recommendations.");

  doc.createParagraph("Methodology").setStyle("Heading1");
  doc.createParagraph("Research approach and methods used.");

  doc.createParagraph("Data Collection").setStyle("Heading2");
  doc.createParagraph("How data was gathered.");

  doc.createParagraph("Analysis Methods").setStyle("Heading2");
  doc.createParagraph("Statistical analysis techniques.");

  doc.createParagraph("Results").setStyle("Heading1");
  doc.createParagraph("Key findings from the research.");

  doc.createParagraph("Quantitative Results").setStyle("Heading2");
  doc.createParagraph("Numerical findings.");

  doc.createParagraph("Qualitative Results").setStyle("Heading2");
  doc.createParagraph("Interview insights.");

  doc.createParagraph("Conclusions").setStyle("Heading1");
  doc.createParagraph("Final conclusions and recommendations.");

  await doc.save(path.join(OUTPUT_DIR, "toc-best-practices.docx"));
  console.log("✓ Created document following TOC best practices");
}

/**
 * Main execution
 */
async function main() {
  console.log("docXMLater - TOC Fix Examples\n");
  console.log("Demonstrating proper TOC usage with new logging system\n");

  // Create output directory
  const fs = require("fs");
  if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  }

  try {
    await example1_BasicTOCWithLogging();
    await example2_TOCSilentMode();
    await example3_TOCWithCustomStyling();
    await example4_MultiLevelTOC();
    await example5_TOCUpdateWorkflow();
    await example6_TOCBestPractices();

    console.log("\n✓ All TOC examples completed successfully!");
    console.log(`Output directory: ${OUTPUT_DIR}`);
    console.log(
      "\nTip: Open any document in Word and press F9 on the TOC to update it."
    );
  } catch (error) {
    console.error("Error running examples:", error);
    process.exit(1);
  }
}

// Run if executed directly
if (require.main === module) {
  main();
}

export {
  example1_BasicTOCWithLogging,
  example2_TOCSilentMode,
  example3_TOCWithCustomStyling,
  example4_MultiLevelTOC,
  example5_TOCUpdateWorkflow,
  example6_TOCBestPractices,
};
