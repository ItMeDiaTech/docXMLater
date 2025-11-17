/**
 * Unlock Tables Utility
 *
 * This script demonstrates how to unlock content-locked tables
 * that prevent editing in Microsoft Word.
 *
 * Common causes:
 * - Tables imported from Google Docs (wrapped in SDTs with contentLocked)
 * - Tables created programmatically with lock flags
 *
 * Usage:
 * ```bash
 * npx ts-node examples/troubleshooting/unlock-tables.ts
 * ```
 */

import { Document } from "../../src/core/Document";
import { StructuredDocumentTag } from "../../src/elements/StructuredDocumentTag";

async function unlockTables() {
  console.log("🔓 Unlocking Tables Utility\n");

  // Load the document with locked tables
  const inputPath = "./Errors.docx";
  console.log(`📂 Loading: ${inputPath}`);

  const doc = await Document.load(inputPath);
  console.log("✅ Document loaded successfully\n");

  // Track unlocked SDTs
  let unlockedCount = 0;
  let sdtCount = 0;

  // Iterate through all body elements
  const bodyElements = doc.getBodyElements();

  for (const element of bodyElements) {
    if (element instanceof StructuredDocumentTag) {
      sdtCount++;

      const isLocked = element.isLocked();
      const tag = element.getTag();
      const id = element.getId();

      console.log(`📦 SDT Found:`);
      console.log(`   Tag: ${tag || "<none>"}`);
      console.log(`   ID: ${id}`);
      console.log(`   Locked: ${isLocked ? "🔒 YES" : "🔓 NO"}`);
      console.log(
        `   Editable: ${element.isContentEditable() ? "✅ YES" : "❌ NO"}`
      );

      if (isLocked) {
        console.log(`   🔧 Unlocking...`);
        element.unlock();
        unlockedCount++;
        console.log(`   ✅ Unlocked successfully`);
      }

      console.log("");
    }
  }

  console.log(`\n📊 Summary:`);
  console.log(`   Total SDTs found: ${sdtCount}`);
  console.log(`   SDTs unlocked: ${unlockedCount}`);
  console.log(`   SDTs already unlocked: ${sdtCount - unlockedCount}\n`);

  // Check for parse errors (ComplexField issues)
  const parseErrors = (doc as any).parser?.getParseErrors() || [];
  if (parseErrors.length > 0) {
    console.log(`⚠️  Parse Warnings: ${parseErrors.length}`);
    for (const err of parseErrors) {
      console.log(`   - ${err.element}: ${err.error.message}`);
    }
    console.log("");
  }

  // Save the unlocked document
  const outputPath = "./Errors_UNLOCKED.docx";
  console.log(`💾 Saving unlocked document to: ${outputPath}`);

  await doc.save(outputPath);

  console.log("✅ Document saved successfully");
  console.log("\n🎉 Done! Tables should now be editable in Word.");
  console.log(`\n📝 Next steps:`);
  console.log(`   1. Open ${outputPath} in Microsoft Word`);
  console.log(`   2. Try clicking in table cells`);
  console.log(`   3. Verify you can type/edit text`);
}

// Run the unlock utility
unlockTables().catch((error) => {
  console.error("❌ Error:", error);
  process.exit(1);
});
