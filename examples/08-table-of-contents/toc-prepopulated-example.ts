/**
 * Table of Contents - Pre-Populated Example
 * Demonstrates how to create TOCs that are automatically populated with heading entries
 * when the document is first opened in Word
 */

import { Document, TOCProperties } from "../../src/index";
import * as path from "path";
import * as fs from "fs";

// Create output directory if it doesn't exist
const outputDir = __dirname;
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

/**
 * Example 1: Simple Pre-Populated TOC
 * Uses the convenience method to create a TOC with entries visible immediately
 */
async function example1_SimplePrePopulated() {
  console.log("Example 1: Simple Pre-Populated TOC...");

  const doc = Document.create();

  // Add document title
  doc
    .createParagraph("Technical Documentation")
    .setStyle("Title")
    .setAlignment("center");

  // Add headings
  doc.createParagraph("Chapter 1: Introduction").setStyle("Heading1");
  doc
    .createParagraph("Overview of the system and its components.")
    .setSpaceAfter(240);

  doc.createParagraph("Section 1.1: Background").setStyle("Heading2");
  doc
    .createParagraph("Historical context and motivation for the project.")
    .setSpaceAfter(240);

  doc.createParagraph("Section 1.2: Objectives").setStyle("Heading2");
  doc.createParagraph("Key goals and success criteria.").setSpaceAfter(480);

  doc.createParagraph("Chapter 2: Architecture").setStyle("Heading1");
  doc
    .createParagraph("System architecture and design patterns.")
    .setSpaceAfter(240);

  doc.createParagraph("Section 2.1: Components").setStyle("Heading2");
  doc
    .createParagraph("Description of major system components.")
    .setSpaceAfter(240);

  doc.createParagraph("Section 2.2: Data Flow").setStyle("Heading2");
  doc.createParagraph("How data moves through the system.").setSpaceAfter(480);

  doc.createParagraph("Chapter 3: Implementation").setStyle("Heading1");
  doc
    .createParagraph("Implementation details and code examples.")
    .setSpaceAfter(240);

  // Create pre-populated TOC using convenience method
  // This will show all heading entries when document is opened
  doc.createPrePopulatedTableOfContents();

  // Save document - TOC will be populated automatically!
  const outputPath = path.join(outputDir, "example1-simple-prepopulated.docx");
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log("  TOC entries are visible when you open the document!");
  console.log(
    '  You can still right-click and "Update Field" if you add more headings.'
  );
}

/**
 * Example 2: Pre-Populated TOC with Custom Options
 * Shows how to customize the TOC while still pre-populating it
 */
async function example2_CustomPrePopulated() {
  console.log("\nExample 2: Custom Pre-Populated TOC...");

  const doc = Document.create();

  doc
    .createParagraph("Research Paper")
    .setStyle("Title")
    .setAlignment("center");

  // Add content with multiple heading levels
  doc.createParagraph("Abstract").setStyle("Heading1");
  doc
    .createParagraph("Summary of research findings and conclusions.")
    .setSpaceAfter(240);

  doc.createParagraph("1. Introduction").setStyle("Heading1");
  doc
    .createParagraph("Background and context for the research.")
    .setSpaceAfter(240);

  doc.createParagraph("1.1 Problem Statement").setStyle("Heading2");
  doc
    .createParagraph("Description of the research problem.")
    .setSpaceAfter(240);

  doc.createParagraph("1.2 Research Questions").setStyle("Heading2");
  doc
    .createParagraph("Key questions this research addresses.")
    .setSpaceAfter(240);

  doc.createParagraph("1.2.1 Primary Questions").setStyle("Heading3");
  doc.createParagraph("Main research questions.").setSpaceAfter(240);

  doc.createParagraph("1.2.2 Secondary Questions").setStyle("Heading3");
  doc.createParagraph("Supporting research questions.").setSpaceAfter(480);

  doc.createParagraph("2. Methodology").setStyle("Heading1");
  doc.createParagraph("Research methods and approach.").setSpaceAfter(240);

  doc.createParagraph("2.1 Data Collection").setStyle("Heading2");
  doc.createParagraph("Methods for gathering data.").setSpaceAfter(240);

  doc.createParagraph("2.2 Analysis").setStyle("Heading2");
  doc.createParagraph("Analytical techniques used.").setSpaceAfter(480);

  doc.createParagraph("3. Results").setStyle("Heading1");
  doc.createParagraph("Findings and observations.").setSpaceAfter(240);

  doc.createParagraph("4. Conclusion").setStyle("Heading1");
  doc.createParagraph("Summary and recommendations.").setSpaceAfter(240);

  // Create pre-populated TOC with custom options
  doc.createPrePopulatedTableOfContents("Table of Contents", {
    levels: 3, // Include Heading1, Heading2, and Heading3
    useHyperlinks: true, // Make entries clickable
    hideInWebLayout: true, // Hide page numbers in web view
  });

  const outputPath = path.join(outputDir, "example2-custom-prepopulated.docx");
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log("  TOC includes 3 levels with hyperlinked entries!");
}

/**
 * Example 3: Manual Control Over Population
 * Shows how to enable auto-population separately from TOC creation
 */
async function example3_ManualPopulation() {
  console.log("\nExample 3: Manual Population Control...");

  const doc = Document.create();

  doc.createParagraph("User Guide").setStyle("Title").setAlignment("center");

  // Add table of contents FIRST (before content)
  doc.createTableOfContents("Contents");

  // Enable auto-population for this document
  doc.setAutoPopulateTOCs(true);

  // Add content AFTER TOC
  doc.createParagraph("Getting Started").setStyle("Heading1");
  doc.createParagraph("How to begin using the application.").setSpaceAfter(240);

  doc.createParagraph("Installation").setStyle("Heading2");
  doc.createParagraph("Step-by-step installation guide.").setSpaceAfter(240);

  doc.createParagraph("Configuration").setStyle("Heading2");
  doc.createParagraph("Setting up your environment.").setSpaceAfter(480);

  doc.createParagraph("User Interface").setStyle("Heading1");
  doc.createParagraph("Overview of the user interface.").setSpaceAfter(240);

  doc.createParagraph("Main Window").setStyle("Heading2");
  doc.createParagraph("Description of the main window.").setSpaceAfter(240);

  doc.createParagraph("Menus and Toolbars").setStyle("Heading2");
  doc
    .createParagraph("Available menus and toolbar options.")
    .setSpaceAfter(480);

  doc.createParagraph("Advanced Features").setStyle("Heading1");
  doc.createParagraph("Power user features and shortcuts.").setSpaceAfter(240);

  const outputPath = path.join(outputDir, "example3-manual-population.docx");
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log("  TOC placed before content, but populated during save!");
}

/**
 * Example 4: Loading and Adding Content
 * Demonstrates auto-population when loading an existing document
 */
async function example4_LoadAndPopulate() {
  console.log("\nExample 4: Load and Populate...");

  // This example would load an existing document
  // For demo purposes, we'll create one first
  const doc = Document.create();

  doc
    .createParagraph("Original Document")
    .setStyle("Title")
    .setAlignment("center");
  doc.createTableOfContents("Table of Contents");

  doc.createParagraph("Chapter 1").setStyle("Heading1");
  doc.createParagraph("Original chapter content.").setSpaceAfter(480);

  doc.createParagraph("Chapter 2").setStyle("Heading1");
  doc.createParagraph("More original content.").setSpaceAfter(240);

  const tempPath = path.join(outputDir, "temp-for-loading.docx");
  await doc.save(tempPath);

  // Now load and add more content
  const loadedDoc = await Document.load(tempPath);

  // Add new headings
  loadedDoc.createParagraph("Chapter 3: New Content").setStyle("Heading1");
  loadedDoc
    .createParagraph("This chapter was added after loading.")
    .setSpaceAfter(240);

  loadedDoc.createParagraph("Section 3.1: Details").setStyle("Heading2");
  loadedDoc.createParagraph("Additional details.").setSpaceAfter(480);

  // Enable auto-population and save
  loadedDoc.setAutoPopulateTOCs(true);

  const outputPath = path.join(outputDir, "example4-loaded-and-populated.docx");
  await loadedDoc.save(outputPath);

  // Clean up temp file
  fs.unlinkSync(tempPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log("  TOC now includes original AND new chapters!");
}

/**
 * Example 5: Comparison - With and Without Pre-Population
 * Creates two documents to show the difference
 */
async function example5_Comparison() {
  console.log("\nExample 5: Comparison...");

  // Document WITHOUT pre-population (default behavior)
  const doc1 = Document.create();
  doc1.createParagraph("Without Pre-Population").setStyle("Title");
  doc1.createTableOfContents(); // Standard TOC

  doc1.createParagraph("Chapter 1").setStyle("Heading1");
  doc1.createParagraph("Content here.").setSpaceAfter(240);

  doc1.createParagraph("Chapter 2").setStyle("Heading1");
  doc1.createParagraph("More content.").setSpaceAfter(240);

  const path1 = path.join(outputDir, "example5-without-prepopulation.docx");
  await doc1.save(path1);
  console.log(`✓ Saved ${path.basename(path1)}`);
  console.log('  Shows: "Right-click to update field."');

  // Document WITH pre-population
  const doc2 = Document.create();
  doc2.createParagraph("With Pre-Population").setStyle("Title");

  doc2.createParagraph("Chapter 1").setStyle("Heading1");
  doc2.createParagraph("Content here.").setSpaceAfter(240);

  doc2.createParagraph("Chapter 2").setStyle("Heading1");
  doc2.createParagraph("More content.").setSpaceAfter(240);

  doc2.createPrePopulatedTableOfContents(); // Pre-populated TOC

  const path2 = path.join(outputDir, "example5-with-prepopulation.docx");
  await doc2.save(path2);
  console.log(`✓ Saved ${path.basename(path2)}`);
  console.log("  Shows: Chapter 1, Chapter 2 (with hyperlinks)");
}

/**
 * Example 6: Complex Document with Multiple Heading Levels
 */
async function example6_ComplexDocument() {
  console.log("\nExample 6: Complex Multi-Level Document...");

  const doc = Document.create();

  doc
    .createParagraph("Enterprise Software Documentation")
    .setStyle("Title")
    .setAlignment("center");

  // Part I
  doc.createParagraph("Part I: Getting Started").setStyle("Heading1");

  doc.createParagraph("Chapter 1: Installation").setStyle("Heading2");
  doc
    .createParagraph("System requirements and installation steps.")
    .setSpaceAfter(240);

  doc.createParagraph("1.1 Prerequisites").setStyle("Heading3");
  doc.createParagraph("Software and hardware requirements.").setSpaceAfter(240);

  doc.createParagraph("1.2 Installation Steps").setStyle("Heading3 ");
  doc.createParagraph("Detailed installation instructions.").setSpaceAfter(480);

  doc.createParagraph("Chapter 2: Configuration").setStyle("Heading2");
  doc.createParagraph("Initial setup and configuration.").setSpaceAfter(240);

  doc.createParagraph("2.1 Basic Setup").setStyle("Heading3");
  doc.createParagraph("Essential configuration options.").setSpaceAfter(240);

  doc.createParagraph("2.2 Advanced Options").setStyle("Heading3");
  doc.createParagraph("Optional advanced configurations.").setSpaceAfter(480);

  // Part II
  doc.createParagraph("Part II: User Guide").setStyle("Heading1");

  doc.createParagraph("Chapter 3: Basic Usage").setStyle("Heading2");
  doc.createParagraph("Common tasks and workflows.").setSpaceAfter(240);

  doc.createParagraph("3.1 Creating Projects").setStyle("Heading3");
  doc.createParagraph("How to create new projects.").setSpaceAfter(240);

  doc.createParagraph("3.2 Managing Files").setStyle("Heading3");
  doc.createParagraph("File organization and management.").setSpaceAfter(480);

  doc.createParagraph("Chapter 4: Advanced Features").setStyle("Heading2");
  doc.createParagraph("Power user features and automation.").setSpaceAfter(240);

  // Part III
  doc.createParagraph("Part III: Reference").setStyle("Heading1");

  doc.createParagraph("Chapter 5: API Reference").setStyle("Heading2");
  doc.createParagraph("Complete API documentation.").setSpaceAfter(240);

  doc.createParagraph("Chapter 6: Troubleshooting").setStyle("Heading2");
  doc.createParagraph("Common issues and solutions.").setSpaceAfter(240);

  // Create detailed pre-populated TOC (4 levels)
  doc.createPrePopulatedTableOfContents("Table of Contents", {
    levels: 4,
    useHyperlinks: true,
    tabLeader: "dot",
  });

  const outputPath = path.join(outputDir, "example6-complex-prepopulated.docx");
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log("  Complex TOC with 4 levels, all entries visible!");
}

/**
 * Example 7: Programmatic Control
 * Shows how to check if auto-population is enabled
 */
async function example7_ProgrammaticControl() {
  console.log("\nExample 7: Programmatic Control...");

  const doc = Document.create();

  doc.createParagraph("Controlled Population Example").setStyle("Title");

  // Add some headings
  doc.createParagraph("Introduction").setStyle("Heading1");
  doc.createParagraph("Methods").setStyle("Heading1");
  doc.createParagraph("Results").setStyle("Heading1");

  // Create TOC
  doc.createTableOfContents();

  // Check current setting
  if (!doc.isAutoPopulateTOCsEnabled()) {
    console.log("  Auto-population is currently disabled");

    // Enable it
    doc.setAutoPopulateTOCs(true);
    console.log("  Auto-population enabled!");
  }

  // Verify it's enabled
  if (doc.isAutoPopulateTOCsEnabled()) {
    console.log("  Confirmed: TOCs will be pre-populated on save");
  }

  const outputPath = path.join(outputDir, "example7-programmatic-control.docx");
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
}

/**
 * Example 8: Web Document TOC
 * Creates a web-optimized TOC with pre-population
 */
async function example8_WebDocumentTOC() {
  console.log("\nExample 8: Web Document TOC...");

  const doc = Document.create();

  doc
    .createParagraph("Online Documentation")
    .setStyle("Title")
    .setAlignment("center");

  // Add navigation-style headings
  doc.createParagraph("Home").setStyle("Heading1");
  doc
    .createParagraph("Welcome to our documentation portal.")
    .setSpaceAfter(240);

  doc.createParagraph("Quick Start").setStyle("Heading1");
  doc.createParagraph("Get up and running in 5 minutes.").setSpaceAfter(240);

  doc.createParagraph("Step 1: Setup").setStyle("Heading2");
  doc.createParagraph("Initial setup instructions.").setSpaceAfter(240);

  doc.createParagraph("Step 2: First Project").setStyle("Heading2");
  doc.createParagraph("Create your first project.").setSpaceAfter(480);

  doc.createParagraph("Tutorials").setStyle("Heading1");
  doc.createParagraph("Step-by-step tutorials.").setSpaceAfter(240);

  doc.createParagraph("API Reference").setStyle("Heading1");
  doc.createParagraph("Complete API documentation.").setSpaceAfter(240);

  doc.createParagraph("FAQ").setStyle("Heading1");
  doc.createParagraph("Frequently asked questions.").setSpaceAfter(240);

  // Web-optimized TOC: hyperlinks, no page numbers
  doc.createPrePopulatedTableOfContents("Navigation", {
    levels: 2,
    useHyperlinks: true,
    showPageNumbers: false,
    hideInWebLayout: true,
  });

  const outputPath = path.join(outputDir, "example8-web-document-toc.docx");
  await doc.save(outputPath);

  console.log(`✓ Saved to ${outputPath}`);
  console.log("  Web-optimized TOC with clickable navigation!");
}

/**
 * Main function to run all examples
 */
async function main() {
  console.log("Running Pre-Populated TOC Examples...\n");

  try {
    await example1_SimplePrePopulated();
    await example2_CustomPrePopulated();
    await example3_ManualPopulation();
    await example4_LoadAndPopulate();
    await example5_Comparison();
    await example6_ComplexDocument();
    await example7_ProgrammaticControl();
    await example8_WebDocumentTOC();

    console.log("\n✓ All examples completed successfully!");
    console.log(`\nOutput files saved to: ${outputDir}`);
    console.log("\n📝 How it works:");
    console.log("   1. TOCs are created with proper field structure");
    console.log("   2. When you save, headings are automatically found");
    console.log("   3. TOC is populated with entries as hyperlinks");
    console.log("   4. Open in Word - entries are already visible!");
    console.log(
      '   5. Right-click "Update Field" still works if you add more headings'
    );
  } catch (error) {
    console.error("Error running examples:", error);
    process.exit(1);
  }
}

// Run examples if executed directly
if (require.main === module) {
  main();
}

export {
  example1_SimplePrePopulated,
  example2_CustomPrePopulated,
  example3_ManualPopulation,
  example4_LoadAndPopulate,
  example5_Comparison,
  example6_ComplexDocument,
  example7_ProgrammaticControl,
  example8_WebDocumentTOC,
};
