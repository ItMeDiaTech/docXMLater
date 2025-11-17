/**
 * Bullet Customization Example
 * Demonstrates how to customize bullet symbols for each level of a bullet list
 */

import { Document, AbstractNumbering, NumberingLevel } from "../src/index";
import * as path from "path";

const OUTPUT_DIR = path.join(__dirname, "output");

/**
 * Example 1: Using Default Alternating Bullets (Closed • and Open ○)
 * Default now alternates: Level 0=•, Level 1=○, Level 2=•, Level 3=○, etc.
 */
async function example1_DefaultAlternatingBullets() {
  console.log("\n=== Example 1: Default Alternating Bullets ===");

  const doc = Document.create();

  // Create bullet list with defaults (• and ○ alternating)
  const bulletListId = doc.createBulletList();

  // Level 0 (•)
  doc.createParagraph("First item").setNumbering(bulletListId, 0);
  doc.createParagraph("Second item").setNumbering(bulletListId, 0);

  // Level 1 (○)
  doc.createParagraph("Nested item A").setNumbering(bulletListId, 1);
  doc.createParagraph("Nested item B").setNumbering(bulletListId, 1);

  // Level 2 (•)
  doc.createParagraph("Deep nested X").setNumbering(bulletListId, 2);
  doc.createParagraph("Deep nested Y").setNumbering(bulletListId, 2);

  await doc.save(path.join(OUTPUT_DIR, "default-alternating-bullets.docx"));
  console.log("✓ Created with default alternating bullets (• ○ • ○ ...)");
}

/**
 * Example 2: Custom Bullet Symbols for All 9 Levels
 * Shows how to specify exactly which symbol to use at each level
 */
async function example2_CustomBulletSymbols() {
  console.log("\n=== Example 2: Custom Bullet Symbols ===");

  const doc = Document.create();

  // Define custom bullets: ■, □, ●, ○, ▪, ▫, ►, ▻, •
  const customBullets = ["■", "□", "●", "○", "▪", "▫", "►", "▻", "•"];
  const bulletListId = doc.createBulletList(9, customBullets);

  // Demonstrate each level
  for (let level = 0; level < 9; level++) {
    doc
      .createParagraph(`Level ${level}: ${customBullets[level]}`)
      .setNumbering(bulletListId, level);
  }

  await doc.save(path.join(OUTPUT_DIR, "custom-bullet-symbols.docx"));
  console.log("✓ Created with custom bullet symbols for all 9 levels");
}

/**
 * Example 3: Pattern-Based Bullets (Repeating Pattern)
 * Uses a 3-symbol pattern that repeats: ●, ○, ▪
 */
async function example3_PatternBullets() {
  console.log("\n=== Example 3: Pattern-Based Bullets ===");

  const doc = Document.create();

  // 3-symbol pattern repeats across all 9 levels
  // Level 0=●, 1=○, 2=▪, 3=●, 4=○, 5=▪, 6=●, 7=○, 8=▪
  const pattern = ["●", "○", "▪"];
  const bulletListId = doc.createBulletList(9, pattern);

  doc.createParagraph("Level 0 (●)").setNumbering(bulletListId, 0);
  doc.createParagraph("Level 1 (○)").setNumbering(bulletListId, 1);
  doc.createParagraph("Level 2 (▪)").setNumbering(bulletListId, 2);
  doc.createParagraph("Level 3 (●)").setNumbering(bulletListId, 3);
  doc.createParagraph("Level 4 (○)").setNumbering(bulletListId, 4);

  await doc.save(path.join(OUTPUT_DIR, "pattern-bullets.docx"));
  console.log("✓ Created with 3-symbol repeating pattern");
}

/**
 * Example 4: Advanced - Modify Existing Numbering Definition
 * Shows how to access and modify bullet symbols in an existing numbering
 */
async function example4_ModifyExistingNumbering() {
  console.log("\n=== Example 4: Modify Existing Numbering ===");

  const doc = Document.create();

  // Create a standard bullet list
  const bulletListId = doc.createBulletList();

  // Get the numbering manager
  const numberingManager = doc.getNumberingManager();

  // Get the instance and its abstract numbering
  const instance = numberingManager.getInstance(bulletListId);
  if (instance) {
    const abstractNum = numberingManager.getAbstractNumbering(
      instance.getAbstractNumId()
    );

    if (abstractNum) {
      // Modify level 0 to use a square bullet
      const level0 = abstractNum.getLevel(0);
      if (level0) {
        level0.setText("■");
        level0.setFont("Arial");
        level0.setFontSize(24); // 12pt
        level0.setBold(true);
        level0.setColor("000000");
      }

      // Modify level 1 to use an arrow
      const level1 = abstractNum.getLevel(1);
      if (level1) {
        level1.setText("➤");
        level1.setFont("Arial");
        level1.setFontSize(24);
        level1.setBold(true);
        level1.setColor("000000");
      }
    }
  }

  doc.createParagraph("Square bullet").setNumbering(bulletListId, 0);
  doc.createParagraph("Arrow bullet").setNumbering(bulletListId, 1);

  await doc.save(path.join(OUTPUT_DIR, "modified-bullets.docx"));
  console.log("✓ Modified existing numbering definition");
}

/**
 * Example 5: Creating Custom Bullet Levels from Scratch
 * Shows low-level control over bullet formatting
 */
async function example5_CustomBulletLevelsFromScratch() {
  console.log("\n=== Example 5: Custom Bullet Levels From Scratch ===");

  const doc = Document.create();
  const numberingManager = doc.getNumberingManager();

  // Get next abstract numbering ID
  const abstractNumId = (numberingManager as any).nextAbstractNumId++;

  // Create custom abstract numbering with fully custom levels
  const levels: NumberingLevel[] = [];

  // Level 0: Red filled circle (●)
  levels.push(
    NumberingLevel.createBulletLevel(0, "●")
      .setFont("Arial")
      .setFontSize(24) // 12pt
      .setBold(true)
      .setColor("FF0000") // Red
  );

  // Level 1: Blue open circle (○)
  levels.push(
    NumberingLevel.createBulletLevel(1, "○")
      .setFont("Arial")
      .setFontSize(24)
      .setBold(true)
      .setColor("0000FF") // Blue
  );

  // Level 2: Green square (■)
  levels.push(
    NumberingLevel.createBulletLevel(2, "■")
      .setFont("Arial")
      .setFontSize(24)
      .setBold(true)
      .setColor("00FF00") // Green
  );

  const abstractNum = AbstractNumbering.create({
    abstractNumId,
    name: "Colored Bullets",
    levels,
  });

  numberingManager.addAbstractNumbering(abstractNum);

  // Create instance
  const numId = (numberingManager as any).nextNumId++;
  const instance = {
    getNumId: () => numId,
    getAbstractNumId: () => abstractNumId,
  };
  numberingManager.addInstance(instance as any);

  doc.createParagraph("Red filled circle").setNumbering(numId, 0);
  doc.createParagraph("Blue open circle").setNumbering(numId, 1);
  doc.createParagraph("Green square").setNumbering(numId, 2);

  await doc.save(path.join(OUTPUT_DIR, "colored-custom-bullets.docx"));
  console.log("✓ Created with colored custom bullet levels");
}

/**
 * Example 6: Reusable Function - Set Custom Bullets for Any Document
 * Create a reusable function you can import in other projects
 */
function setCustomBulletDefaults(
  doc: Document,
  bulletSymbols: string[]
): number {
  /**
   * Sets custom default bullet symbols for a document
   *
   * @param doc - The Document instance
   * @param bulletSymbols - Array of bullet symbols to use
   * @returns The numId of the created bullet list
   *
   * @example
   * // In your project:
   * import { setCustomBulletDefaults } from './bullet-config';
   *
   * const doc = Document.create();
   * const myBullets = ['★', '☆', '✦', '✧'];
   * const bulletListId = setCustomBulletDefaults(doc, myBullets);
   *
   * doc.createParagraph('Star bullet').setNumbering(bulletListId, 0);
   * await doc.save('custom-bullets.docx');
   */

  return doc.createBulletList(9, bulletSymbols);
}

async function example6_ReusableFunction() {
  console.log("\n=== Example 6: Reusable Function ===");

  const doc = Document.create();

  // Use custom symbols: stars and hearts
  const myBullets = ["★", "☆", "♥", "♡"];
  const bulletListId = setCustomBulletDefaults(doc, myBullets);

  doc.createParagraph("Filled star").setNumbering(bulletListId, 0);
  doc.createParagraph("Open star").setNumbering(bulletListId, 1);
  doc.createParagraph("Filled heart").setNumbering(bulletListId, 2);
  doc.createParagraph("Open heart").setNumbering(bulletListId, 3);

  await doc.save(path.join(OUTPUT_DIR, "reusable-function-bullets.docx"));
  console.log("✓ Created using reusable function");
}

/**
 * Main execution
 */
async function main() {
  console.log("docXMLater - Bullet Customization Examples\n");

  // Create output directory
  const fs = require("fs");
  if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  }

  try {
    await example1_DefaultAlternatingBullets();
    await example2_CustomBulletSymbols();
    await example3_PatternBullets();
    await example4_ModifyExistingNumbering();
    await example5_CustomBulletLevelsFromScratch();
    await example6_ReusableFunction();

    console.log("\n✓ All bullet customization examples completed!");
    console.log(`Output directory: ${OUTPUT_DIR}`);
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
  example1_DefaultAlternatingBullets,
  example2_CustomBulletSymbols,
  example3_PatternBullets,
  example4_ModifyExistingNumbering,
  example5_CustomBulletLevelsFromScratch,
  example6_ReusableFunction,
  setCustomBulletDefaults,
};
