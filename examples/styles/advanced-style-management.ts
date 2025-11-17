/**
 * Advanced Style Management Examples
 *
 * Demonstrates the new StylesManager enhancements for searching, filtering,
 * validation, cleanup, inheritance analysis, and export/import operations.
 * These examples show real-world usage scenarios for managing document styles.
 */

import { Document, Style } from "../..";

/**
 * Demonstrates style search and filtering capabilities
 */
async function demonstrateSearchAndFilter() {
  console.log("=== Search & Filter Examples ===\n");

  const doc = Document.create({
    properties: {
      title: "Style Search & Filter Demo",
      creator: "DocXML Advanced Examples",
    },
  });

  // Create a variety of styles for demonstration
  const stylesManager = doc.getStylesManager();

  // Add custom styles with different fonts and names
  const arialStyle = Style.create({
    styleId: "ArialBody",
    name: "Arial Body Text",
    type: "paragraph",
    basedOn: "Normal",
    customStyle: true,
    runFormatting: { font: "Arial", size: 11 },
  });
  stylesManager.addStyle(arialStyle);

  const headingStyle = Style.create({
    styleId: "CustomHeading",
    name: "Custom Heading Style",
    type: "paragraph",
    basedOn: "Normal",
    customStyle: true,
    runFormatting: { font: "Arial", size: 14, bold: true },
  });
  stylesManager.addStyle(headingStyle);

  const codeStyle = Style.create({
    styleId: "CodeSnippet",
    name: "Code Snippet",
    type: "paragraph",
    basedOn: "Normal",
    customStyle: true,
    runFormatting: { font: "Consolas", size: 10 },
  });
  stylesManager.addStyle(codeStyle);

  const noteStyle = Style.create({
    styleId: "NoteStyle",
    name: "Important Note",
    type: "paragraph",
    basedOn: "Normal",
    customStyle: true,
    runFormatting: { font: "Arial", size: 10, italic: true },
  });
  stylesManager.addStyle(noteStyle);

  // Example 1: searchByName() - Find styles by name
  console.log("1. searchByName() Examples:");
  console.log('   Finding styles containing "Heading":');
  const headingStyles = stylesManager.searchByName("Heading");
  console.log(
    `   Found ${headingStyles.length} styles: ${headingStyles
      .map((s) => s.getName())
      .join(", ")}`
  );

  console.log('   Finding styles containing "Code":');
  const codeStyles = stylesManager.searchByName("Code");
  console.log(
    `   Found ${codeStyles.length} styles: ${codeStyles
      .map((s) => s.getName())
      .join(", ")}`
  );

  console.log('   Case-insensitive search for "note":');
  const noteStyles = stylesManager.searchByName("note");
  console.log(
    `   Found ${noteStyles.length} styles: ${noteStyles
      .map((s) => s.getName())
      .join(", ")}`
  );
  console.log();

  // Example 2: findByFont() - Find styles using specific font
  console.log("2. findByFont() Examples:");
  console.log("   Finding styles using Arial font:");
  const arialStyles = stylesManager.findByFont("Arial");
  console.log(
    `   Found ${arialStyles.length} styles: ${arialStyles
      .map((s) => s.getName())
      .join(", ")}`
  );

  console.log("   Finding styles using Consolas font:");
  const consolasStyles = stylesManager.findByFont("Consolas");
  console.log(
    `   Found ${consolasStyles.length} styles: ${consolasStyles
      .map((s) => s.getName())
      .join(", ")}`
  );

  console.log("   Finding styles using Calibri (built-in Normal style):");
  const calibriStyles = stylesManager.findByFont("Calibri");
  console.log(
    `   Found ${calibriStyles.length} styles: ${calibriStyles
      .map((s) => s.getName())
      .join(", ")}`
  );
  console.log();

  // Example 3: findStyles() - Generic predicate-based search
  console.log("3. findStyles() Examples:");
  console.log("   Finding all paragraph styles:");
  const paragraphStyles = stylesManager.findStyles(
    (s) => s.getType() === "paragraph"
  );
  console.log(`   Found ${paragraphStyles.length} paragraph styles`);

  console.log("   Finding styles with custom formatting:");
  const customFormattedStyles = stylesManager.findStyles(
    (s) => s.getProperties().customStyle === true
  );
  console.log(`   Found ${customFormattedStyles.length} custom styles`);

  console.log("   Finding styles with bold formatting:");
  const boldStyles = stylesManager.findStyles(
    (s) => s.getRunFormatting()?.bold === true
  );
  console.log(
    `   Found ${boldStyles.length} bold styles: ${boldStyles
      .map((s) => s.getName())
      .join(", ")}`
  );

  console.log("   Finding styles with size > 12pt:");
  const largeStyles = stylesManager.findStyles(
    (s) => (s.getRunFormatting()?.size || 0) > 12
  );
  console.log(
    `   Found ${largeStyles.length} large styles: ${largeStyles
      .map((s) => s.getName())
      .join(", ")}`
  );

  console.log("   Finding styles based on Normal:");
  const normalBasedStyles = stylesManager.findStyles(
    (s) => s.getProperties().basedOn === "Normal"
  );
  console.log(
    `   Found ${
      normalBasedStyles.length
    } styles based on Normal: ${normalBasedStyles
      .map((s) => s.getName())
      .join(", ")}`
  );
  console.log();

  // Create document content demonstrating the found styles
  doc.createParagraph("Style Search & Filter Demonstration").setStyle("Title");
  doc
    .createParagraph(
      "This document demonstrates the StylesManager search and filter capabilities."
    )
    .setStyle("Normal");
  doc.createParagraph();

  doc.createParagraph("Found Heading Styles").setStyle("Heading1");
  headingStyles.forEach((style) => {
    doc
      .createParagraph(`• ${style.getName()} (${style.getStyleId()})`)
      .setStyle("Normal");
  });
  doc.createParagraph();

  doc.createParagraph("Found Arial Font Styles").setStyle("Heading1");
  arialStyles.forEach((style) => {
    doc
      .createParagraph(
        `• ${style.getName()} - Font: ${style.getRunFormatting()?.font}`
      )
      .setStyle("Normal");
  });
  doc.createParagraph();

  doc.createParagraph("Found Bold Styles").setStyle("Heading1");
  boldStyles.forEach((style) => {
    doc
      .createParagraph(
        `• ${style.getName()} - Size: ${style.getRunFormatting()?.size}pt`
      )
      .setStyle("Normal");
  });

  return doc;
}

/**
 * Demonstrates style validation and cleanup capabilities
 */
async function demonstrateValidationAndCleanup() {
  console.log("=== Validation & Cleanup Examples ===\n");

  const doc = Document.create({
    properties: {
      title: "Style Validation & Cleanup Demo",
      creator: "DocXML Advanced Examples",
    },
  });

  const stylesManager = doc.getStylesManager();

  // Create styles for the document
  const usedStyle = Style.create({
    styleId: "UsedStyle",
    name: "Used Style",
    type: "paragraph",
    basedOn: "Normal",
    customStyle: true,
  });
  stylesManager.addStyle(usedStyle);

  const unusedStyle = Style.create({
    styleId: "UnusedStyle",
    name: "Unused Style",
    type: "paragraph",
    basedOn: "Normal",
    customStyle: true,
  });
  stylesManager.addStyle(unusedStyle);

  const brokenRefStyle = Style.create({
    styleId: "BrokenRefStyle",
    name: "Broken Reference Style",
    type: "paragraph",
    basedOn: "NonExistentStyle", // This will create a broken reference
    customStyle: true,
  });
  stylesManager.addStyle(brokenRefStyle);

  // Add content using some styles but not others
  doc.createParagraph("Document with Mixed Style Usage").setStyle("Title");
  doc
    .createParagraph("This paragraph uses a style that will be referenced.")
    .setStyle("UsedStyle");
  doc.createParagraph("This paragraph uses Normal style.").setStyle("Normal");
  doc
    .createParagraph("Another paragraph with UsedStyle.")
    .setStyle("UsedStyle");

  // Example 4: findUnusedStyles() - Detect orphaned styles
  console.log("4. findUnusedStyles() Example:");
  const unused = stylesManager.findUnusedStyles(doc.getAllParagraphs());
  console.log(`   Found ${unused.length} unused styles: ${unused.join(", ")}`);
  console.log();

  // Example 5: cleanupUnusedStyles() - Remove orphaned styles
  console.log("5. cleanupUnusedStyles() Example:");
  const removedCount = stylesManager.cleanupUnusedStyles(
    doc.getAllParagraphs()
  );
  console.log(`   Removed ${removedCount} unused styles`);
  console.log("   Remaining styles after cleanup:");
  stylesManager.getAllStyles().forEach((style) => {
    console.log(`   - ${style.getStyleId()}: ${style.getName()}`);
  });
  console.log();

  // Example 6: validateStyleReferences() - Check broken references
  console.log("6. validateStyleReferences() Example:");

  // First, let's add back the broken reference style to demonstrate validation
  stylesManager.addStyle(brokenRefStyle);

  const validation = stylesManager.validateStyleReferences();
  console.log(
    `   Validation result: ${validation.valid ? "VALID" : "INVALID"}`
  );
  if (!validation.valid) {
    console.log(
      `   Broken references (${validation.brokenReferences.length}):`
    );
    validation.brokenReferences.forEach((ref) => {
      console.log(`   - ${ref.styleId} references non-existent ${ref.basedOn}`);
    });

    if (validation.circularReferences.length > 0) {
      console.log(
        `   Circular references (${validation.circularReferences.length}):`
      );
      validation.circularReferences.forEach((chain) => {
        console.log(`   - ${chain.join(" → ")}`);
      });
    }
  }
  console.log();

  // Create document content showing validation results
  doc.createParagraph("Style Validation Results").setStyle("Heading1");
  doc
    .createParagraph(
      `Validation Status: ${
        validation.valid ? "All styles valid" : "Issues found"
      }`
    )
    .setStyle("Normal");

  if (!validation.valid) {
    if (validation.brokenReferences.length > 0) {
      doc.createParagraph("Broken References:").setStyle("Normal");
      validation.brokenReferences.forEach((ref) => {
        doc
          .createParagraph(`• ${ref.styleId} → ${ref.basedOn} (does not exist)`)
          .setStyle("Normal");
      });
    }

    if (validation.circularReferences.length > 0) {
      doc.createParagraph("Circular References:").setStyle("Normal");
      validation.circularReferences.forEach((chain) => {
        doc.createParagraph(`• ${chain.join(" → ")}`).setStyle("Normal");
      });
    }
  }

  doc.createParagraph("Cleanup Summary").setStyle("Heading1");
  doc
    .createParagraph(`Removed ${removedCount} unused styles during cleanup.`)
    .setStyle("Normal");

  return doc;
}

/**
 * Demonstrates style inheritance analysis capabilities
 */
async function demonstrateInheritanceAnalysis() {
  console.log("=== Inheritance Analysis Examples ===\n");

  const doc = Document.create({
    properties: {
      title: "Style Inheritance Analysis Demo",
      creator: "DocXML Advanced Examples",
    },
  });

  const stylesManager = doc.getStylesManager();

  // Create a style hierarchy for demonstration
  const baseStyle = Style.create({
    styleId: "BaseStyle",
    name: "Base Style",
    type: "paragraph",
    basedOn: "Normal",
    customStyle: true,
    runFormatting: { font: "Arial", size: 11 },
  });
  stylesManager.addStyle(baseStyle);

  const headingBase = Style.create({
    styleId: "HeadingBase",
    name: "Heading Base",
    type: "paragraph",
    basedOn: "BaseStyle",
    customStyle: true,
    runFormatting: { bold: true, color: "000000" },
  });
  stylesManager.addStyle(headingBase);

  const heading1 = Style.create({
    styleId: "Heading1Custom",
    name: "Custom Heading 1",
    type: "paragraph",
    basedOn: "HeadingBase",
    customStyle: true,
    runFormatting: { size: 16 },
  });
  stylesManager.addStyle(heading1);

  const heading2 = Style.create({
    styleId: "Heading2Custom",
    name: "Custom Heading 2",
    type: "paragraph",
    basedOn: "HeadingBase",
    customStyle: true,
    runFormatting: { size: 14, italic: true },
  });
  stylesManager.addStyle(heading2);

  const bodyStyle = Style.create({
    styleId: "BodyStyle",
    name: "Body Style",
    type: "paragraph",
    basedOn: "BaseStyle",
    customStyle: true,
    runFormatting: { color: "333333" },
  });
  stylesManager.addStyle(bodyStyle);

  // Example 7: getInheritanceChain() - Full inheritance path
  console.log("7. getInheritanceChain() Examples:");
  console.log("   Inheritance chain for Heading1Custom:");
  try {
    const h1Chain = stylesManager.getInheritanceChain("Heading1Custom");
    h1Chain.forEach((style, index) => {
      const indent = "  ".repeat(index);
      console.log(`   ${indent}${style.getStyleId()}: ${style.getName()}`);
    });
  } catch (error) {
    console.log(
      `   Error: ${error instanceof Error ? error.message : "Unknown error"}`
    );
  }

  console.log("   Inheritance chain for BodyStyle:");
  try {
    const bodyChain = stylesManager.getInheritanceChain("BodyStyle");
    bodyChain.forEach((style, index) => {
      const indent = "  ".repeat(index);
      console.log(`   ${indent}${style.getStyleId()}: ${style.getName()}`);
    });
  } catch (error) {
    console.log(
      `   Error: ${error instanceof Error ? error.message : "Unknown error"}`
    );
  }

  console.log("   Inheritance chain for non-existent style:");
  try {
    stylesManager.getInheritanceChain("NonExistentStyle");
  } catch (error) {
    console.log(
      `   Expected error: ${
        error instanceof Error ? error.message : "Unknown error"
      }`
    );
  }
  console.log();

  // Example 8: getDerivedStyles() - Children of a style
  console.log("8. getDerivedStyles() Examples:");
  console.log("   Styles derived from BaseStyle:");
  const baseChildren = stylesManager.getDerivedStyles("BaseStyle");
  console.log(
    `   Found ${baseChildren.length} child styles: ${baseChildren
      .map((s) => s.getName())
      .join(", ")}`
  );

  console.log("   Styles derived from HeadingBase:");
  const headingChildren = stylesManager.getDerivedStyles("HeadingBase");
  console.log(
    `   Found ${headingChildren.length} child styles: ${headingChildren
      .map((s) => s.getName())
      .join(", ")}`
  );

  console.log("   Styles derived from Normal (built-in):");
  const normalChildren = stylesManager.getDerivedStyles("Normal");
  console.log(
    `   Found ${normalChildren.length} child styles: ${normalChildren
      .map((s) => s.getName())
      .join(", ")}`
  );

  console.log("   Styles derived from non-existent style:");
  const noneChildren = stylesManager.getDerivedStyles("NonExistentStyle");
  console.log(`   Found ${noneChildren.length} child styles`);
  console.log();

  // Create document content showing inheritance analysis
  doc.createParagraph("Style Inheritance Analysis").setStyle("Title");
  doc
    .createParagraph(
      "This document demonstrates inheritance chain analysis and derived style discovery."
    )
    .setStyle("Normal");
  doc.createParagraph();

  doc.createParagraph("Style Hierarchy").setStyle("Heading1");
  doc.createParagraph("Normal (built-in)").setStyle("Normal");
  doc.createParagraph("  └─ BaseStyle (Arial, 11pt)").setStyle("Normal");
  doc.createParagraph("      ├─ HeadingBase (bold, black)").setStyle("Normal");
  doc.createParagraph("      │   ├─ Heading1Custom (16pt)").setStyle("Normal");
  doc
    .createParagraph("      │   └─ Heading2Custom (14pt, italic)")
    .setStyle("Normal");
  doc.createParagraph("      └─ BodyStyle (dark gray)").setStyle("Normal");
  doc.createParagraph();

  doc.createParagraph("Inheritance Chains").setStyle("Heading1");
  doc.createParagraph("Heading1Custom Chain:").setStyle("Normal");
  try {
    const h1Chain = stylesManager.getInheritanceChain("Heading1Custom");
    h1Chain.forEach((style) => {
      doc
        .createParagraph(`• ${style.getName()} (${style.getStyleId()})`)
        .setStyle("Normal");
    });
  } catch (error) {
    doc
      .createParagraph(
        `Error: ${error instanceof Error ? error.message : "Unknown"}`
      )
      .setStyle("Normal");
  }

  doc.createParagraph("Derived Styles Analysis").setStyle("Heading1");
  doc
    .createParagraph(`Styles based on BaseStyle: ${baseChildren.length}`)
    .setStyle("Normal");
  baseChildren.forEach((style) => {
    doc.createParagraph(`• ${style.getName()}`).setStyle("Normal");
  });

  doc
    .createParagraph(`Styles based on HeadingBase: ${headingChildren.length}`)
    .setStyle("Normal");
  headingChildren.forEach((style) => {
    doc.createParagraph(`• ${style.getName()}`).setStyle("Normal");
  });

  return doc;
}

/**
 * Demonstrates style export and import capabilities
 */
async function demonstrateExportImport() {
  console.log("=== Export/Import Examples ===\n");

  const doc = Document.create({
    properties: {
      title: "Style Export/Import Demo",
      creator: "DocXML Advanced Examples",
    },
  });

  const stylesManager = doc.getStylesManager();

  // Create styles to export
  const exportStyle1 = Style.create({
    styleId: "ExportStyle1",
    name: "Style for Export 1",
    type: "paragraph",
    basedOn: "Normal",
    customStyle: true,
    runFormatting: { font: "Arial", size: 12, bold: true, color: "FF0000" },
    paragraphFormatting: { alignment: "center" },
  });
  stylesManager.addStyle(exportStyle1);

  const exportStyle2 = Style.create({
    styleId: "ExportStyle2",
    name: "Style for Export 2",
    type: "character",
    basedOn: "Normal",
    customStyle: true,
    runFormatting: { font: "Consolas", size: 10, color: "008000" },
  });
  stylesManager.addStyle(exportStyle2);

  // Example 9: exportStyle() / importStyle() - Single style JSON
  console.log("9. exportStyle() / importStyle() Examples:");
  console.log("   Exporting ExportStyle1:");
  try {
    const exportedJson = stylesManager.exportStyle("ExportStyle1");
    console.log(
      "   Exported JSON (first 100 chars):",
      exportedJson.substring(0, 100) + "..."
    );

    // Create a new document and import the style
    const importDoc = Document.create();
    const importedStyle = importDoc
      .getStylesManager()
      .importStyle(exportedJson);
    console.log(
      `   Imported style: ${importedStyle.getName()} (${importedStyle.getStyleId()})`
    );
  } catch (error) {
    console.log(
      `   Error: ${error instanceof Error ? error.message : "Unknown error"}`
    );
  }

  console.log("   Attempting to export non-existent style:");
  try {
    stylesManager.exportStyle("NonExistentStyle");
  } catch (error) {
    console.log(
      `   Expected error: ${
        error instanceof Error ? error.message : "Unknown error"
      }`
    );
  }
  console.log();

  // Example 10: exportAllStyles() / importStyles() - Bulk operations
  console.log("10. exportAllStyles() / importStyles() Examples:");
  console.log("    Exporting all styles:");
  const allStylesJson = stylesManager.exportAllStyles();
  console.log(
    `    Exported ${stylesManager.getStyleCount()} styles as JSON (${
      allStylesJson.length
    } chars)`
  );

  // Create a new document and import all styles
  const bulkImportDoc = Document.create();
  try {
    const importedStyles = bulkImportDoc
      .getStylesManager()
      .importStyles(allStylesJson);
    console.log(`    Imported ${importedStyles.length} styles:`);
    importedStyles.forEach((style) => {
      console.log(`    - ${style.getName()} (${style.getStyleId()})`);
    });
  } catch (error) {
    console.log(
      `    Error: ${error instanceof Error ? error.message : "Unknown error"}`
    );
  }

  console.log("    Attempting to import invalid JSON:");
  try {
    stylesManager.importStyles("invalid json");
  } catch (error) {
    console.log(
      `    Expected error: ${
        error instanceof Error ? error.message : "Unknown error"
      }`
    );
  }

  console.log("    Attempting to import non-array JSON:");
  try {
    stylesManager.importStyles('"not an array"');
  } catch (error) {
    console.log(
      `    Expected error: ${
        error instanceof Error ? error.message : "Unknown error"
      }`
    );
  }
  console.log();

  // Create document content showing export/import results
  doc.createParagraph("Style Export/Import Demonstration").setStyle("Title");
  doc
    .createParagraph(
      "This document demonstrates the export and import capabilities of the StylesManager."
    )
    .setStyle("Normal");
  doc.createParagraph();

  doc.createParagraph("Single Style Export/Import").setStyle("Heading1");
  doc
    .createParagraph(
      "ExportStyle1 was exported to JSON and successfully imported into a new document."
    )
    .setStyle("Normal");
  doc
    .createParagraph(
      "This allows styles to be shared between documents or stored externally."
    )
    .setStyle("Normal");
  doc.createParagraph();

  doc.createParagraph("Bulk Export/Import").setStyle("Heading1");
  doc
    .createParagraph(
      `All ${stylesManager.getStyleCount()} styles were exported as a single JSON array.`
    )
    .setStyle("Normal");
  doc
    .createParagraph(
      "This is useful for backing up style collections or transferring them between projects."
    )
    .setStyle("Normal");
  doc.createParagraph();

  doc.createParagraph("Error Handling").setStyle("Heading1");
  doc
    .createParagraph(
      "The API properly handles errors for non-existent styles and invalid JSON."
    )
    .setStyle("Normal");
  doc
    .createParagraph(
      "This ensures robust operation in production environments."
    )
    .setStyle("Normal");

  return doc;
}

/**
 * Main demonstration function that runs all examples
 */
async function demonstrateAdvancedStyleManagement() {
  console.log("🚀 Advanced Style Management Examples\n");
  console.log(
    "This demonstration shows the new StylesManager enhancements for:"
  );
  console.log("• Style searching and filtering");
  console.log("• Validation and cleanup");
  console.log("• Inheritance analysis");
  console.log("• Export and import operations\n");

  try {
    // Run all demonstrations
    const searchDoc = await demonstrateSearchAndFilter();
    const validationDoc = await demonstrateValidationAndCleanup();
    const inheritanceDoc = await demonstrateInheritanceAnalysis();
    const exportDoc = await demonstrateExportImport();

    // Save all demonstration documents
    const docs = [
      { doc: searchDoc, name: "advanced-style-search-filter.docx" },
      { doc: validationDoc, name: "advanced-style-validation-cleanup.docx" },
      { doc: inheritanceDoc, name: "advanced-style-inheritance-analysis.docx" },
      { doc: exportDoc, name: "advanced-style-export-import.docx" },
    ];

    for (const { doc, name } of docs) {
      await doc.save(name);
      console.log(`✓ Created ${name}`);
    }

    console.log(
      "\n🎉 All advanced style management examples completed successfully!"
    );
    console.log(
      "Open the generated DOCX files to see the results of each demonstration."
    );
    console.log("\nKey Features Demonstrated:");
    console.log("• searchByName() - Find styles by partial name matching");
    console.log("• findByFont() - Locate styles using specific fonts");
    console.log("• findStyles() - Generic predicate-based filtering");
    console.log("• findUnusedStyles() - Detect orphaned styles");
    console.log("• cleanupUnusedStyles() - Remove unused styles automatically");
    console.log("• validateStyleReferences() - Check for broken inheritance");
    console.log("• getInheritanceChain() - Analyze complete style hierarchies");
    console.log("• getDerivedStyles() - Find all children of a base style");
    console.log(
      "• exportStyle() / importStyle() - Single style JSON operations"
    );
    console.log("• exportAllStyles() / importStyles() - Bulk style management");
  } catch (error) {
    console.error("❌ Error during demonstration:", error);
    throw error;
  }
}

// Run the complete demonstration
demonstrateAdvancedStyleManagement().catch(console.error);
