/**
 * ListNormalizer - Core list normalization for docxmlater
 *
 * Detects typed list prefixes and converts them to proper Word list formatting.
 * Integrates with NumberingManager for numId resolution.
 */

import type { Paragraph } from "../elements/Paragraph";
import type { Run } from "../elements/Run";
import type { Table } from "../elements/Table";
import type { TableCell } from "../elements/TableCell";
import type { NumberingManager } from "../formatting/NumberingManager";
import type {
  ListCategory,
  ListAnalysis,
  ListNormalizationOptions,
  ListNormalizationReport,
  NumberFormat,
  IndentationLevel,
} from "../types/list-types";
import {
  detectListType,
  detectTypedPrefix,
  getListCategoryFromFormat,
  inferLevelFromRelativeIndentation,
} from "../utils/list-detection";
import { defaultLogger } from "../utils/logger";
import { isRun } from "../elements/Paragraph";

// =============================================================================
// INDENTATION SETTINGS HELPERS
// =============================================================================

/** Helper to convert inches to twips (1 inch = 1440 twips) */
function inchesToTwips(inches: number): number {
  return Math.round(inches * 1440);
}

/**
 * Apply user's indentation settings to an abstract numbering definition.
 */
function applyIndentationSettings(
  abstractNum: ReturnType<NumberingManager["getAbstractNumbering"]>,
  indentationLevels: IndentationLevel[],
  isBulletList: boolean
): void {
  if (!abstractNum || !indentationLevels || indentationLevels.length === 0) return;

  for (const levelConfig of indentationLevels) {
    const level = abstractNum.getLevel(levelConfig.level);
    if (level) {
      const textIndentTwips = inchesToTwips(levelConfig.textIndent);
      const symbolIndentTwips = inchesToTwips(levelConfig.symbolIndent);
      const hangingTwips = textIndentTwips - symbolIndentTwips;

      level.setLeftIndent(textIndentTwips);
      level.setHangingIndent(hangingTwips);

      if (isBulletList && levelConfig.bulletChar) {
        level.setText(levelConfig.bulletChar);
      }
      if (!isBulletList && levelConfig.numberedFormat) {
        level.setFormat(levelConfig.numberedFormat as NumberFormat);
      }
    }
  }
}

// =============================================================================
// ANALYSIS FUNCTIONS
// =============================================================================

/** Internal type for analyzed paragraph data */
interface AnalyzedParagraph {
  paragraph: Paragraph;
  text: string;
  detection: ReturnType<typeof detectListType>;
}

/**
 * Determine majority category using OVERALL counts.
 * Counts ALL list items equally (Word lists + typed prefixes).
 * NUMBERED wins ties (business document standard).
 */
function determineMajorityCategory(analyzed: AnalyzedParagraph[]): ListCategory {
  let bulletCount = 0;
  let numberedCount = 0;

  for (const item of analyzed) {
    // Count BOTH Word lists AND typed prefixes equally
    if (item.detection.category === "bullet") {
      bulletCount++;
    } else if (item.detection.category === "numbered") {
      numberedCount++;
    }
  }

  // No list items at all
  if (bulletCount === 0 && numberedCount === 0) return "none";

  // NUMBERED wins ties (business document standard)
  // Bullets only win if strictly more bullets than numbers
  return numberedCount >= bulletCount ? "numbered" : "bullet";
}

/**
 * Analyze all paragraphs in a cell for list properties.
 */
export function analyzeCellLists(
  cell: TableCell,
  numberingManager?: NumberingManager
): ListAnalysis {
  const paragraphs = cell.getParagraphs();

  const analyzed: AnalyzedParagraph[] = paragraphs.map((p) => ({
    paragraph: p,
    text: p.getText(),
    detection: detectListType(p),
  }));

  // CRITICAL FIX (Round 4): Refine Word list categories using NumberingManager
  // detectListType() defaults ALL Word lists to "numbered", but we need to
  // look up the actual format to correctly identify bullets vs numbers
  if (numberingManager) {
    for (const item of analyzed) {
      if (item.detection.isWordList && item.detection.numId !== null) {
        // Look up the actual format from numbering.xml
        const instance = numberingManager.getInstance(item.detection.numId);
        if (instance) {
          const abstractNum = numberingManager.getAbstractNumbering(
            instance.getAbstractNumId()
          );
          if (abstractNum) {
            const level = abstractNum.getLevel(item.detection.ilvl ?? 0);
            if (level) {
              const format = level.getFormat();
              // Refine the category based on actual format
              item.detection.category = getListCategoryFromFormat(format);
            }
          }
        }
      }
    }
  }

  // Level is now determined by FORMAT in detectListType()
  // decimal=0, lowerLetter=1, lowerRoman=2

  // Count by category
  const counts = { numbered: 0, bullet: 0, none: 0 };
  let hasTypedLists = false;
  let hasWordLists = false;

  for (const item of analyzed) {
    const cat = item.detection.category;
    counts[cat]++;

    if (!item.detection.isWordList && item.detection.typedPrefix) {
      hasTypedLists = true;
    }
    if (item.detection.isWordList) {
      hasWordLists = true;
    }
  }

  // Determine majority using OVERALL counts (Word + typed equally)
  const majorityCategory = determineMajorityCategory(analyzed);

  // Determine if normalization is needed:
  // - Has typed prefixes that need converting, OR
  // - Has mixed categories (bullets AND numbers) that need unifying
  const hasMixedCategories = counts.numbered > 0 && counts.bullet > 0;
  const needsNormalization = hasTypedLists || hasMixedCategories;

  return {
    paragraphs: analyzed,
    hasTypedLists,
    hasWordLists,
    hasMixedCategories,
    majorityCategory,
    counts,
    recommendedAction: needsNormalization ? "normalize" : "none",
  };
}

/**
 * Analyze lists in an entire table.
 * Returns analysis per cell.
 */
export function analyzeTableLists(
  table: Table
): Map<TableCell, ListAnalysis> {
  const results = new Map<TableCell, ListAnalysis>();

  for (const row of table.getRows()) {
    for (const cell of row.getCells()) {
      results.set(cell, analyzeCellLists(cell));
    }
  }

  return results;
}

// =============================================================================
// NORMALIZATION FUNCTIONS
// =============================================================================

/**
 * Strip typed prefix from paragraph text.
 * Handles prefixes that may be split across multiple runs.
 * Also trims leading whitespace from the remaining content.
 */
export function stripTypedPrefix(paragraph: Paragraph, prefix: string): void {
  const content = paragraph.getContent();
  let remainingPrefix = prefix;
  let prefixFullyStripped = false;

  for (const item of content) {
    if (isRun(item)) {
      const run = item as Run;
      const text = run.getText();

      if (!prefixFullyStripped && remainingPrefix.length > 0) {
        if (text.length <= remainingPrefix.length) {
          // Entire run is part of prefix
          if (remainingPrefix.startsWith(text)) {
            remainingPrefix = remainingPrefix.substring(text.length);
            run.setText(""); // Clear this run
            if (remainingPrefix.length === 0) {
              prefixFullyStripped = true;
            }
          }
        } else {
          // Partial match - strip prefix portion
          if (text.startsWith(remainingPrefix)) {
            run.setText(text.substring(remainingPrefix.length).trimStart());
            prefixFullyStripped = true;
          }
        }
      } else if (prefixFullyStripped) {
        // After stripping prefix, trim leading whitespace from first non-empty run
        const currentText = run.getText();
        if (currentText.length > 0) {
          const trimmed = currentText.trimStart();
          if (trimmed !== currentText) {
            run.setText(trimmed);
          }
          break; // Only trim the first run after the prefix
        }
      }
    }
  }
}

/**
 * Normalize all lists in a cell to consistent formatting.
 * KEY BEHAVIORS:
 * - ONE list type per cell - no mixing bullets and numbers
 * - Format determines level: decimal=0, letter=1, roman=2
 * - Word lists that don't match majority are converted
 * - Non-list items are NEVER touched
 * - User indentation settings are applied when provided
 */
export function normalizeListsInCell(
  cell: TableCell,
  options: ListNormalizationOptions,
  numberingManager: NumberingManager
): ListNormalizationReport {
  const analysis = analyzeCellLists(cell, numberingManager);
  const majorityCategory = analysis.majorityCategory;
  const report: ListNormalizationReport = {
    normalized: 0,
    skipped: 0,
    errors: [],
    appliedCategory: majorityCategory,
    details: [],
  };

  // Handle cells that don't need category normalization but may need indentation fixes
  if (analysis.recommendedAction === "none") {
    // Even if no normalization needed, still apply user indentation settings to Word lists
    if (options?.indentationLevels?.length && analysis.hasWordLists) {
      // Create numId maps for applying correct settings
      const correctedNumIdByLevel = new Map<string, number>();
      const getCorrectedNumId = (level: number, isBullet: boolean): number => {
        const key = `${isBullet ? 'b' : 'n'}-${level}`;
        if (!correctedNumIdByLevel.has(key)) {
          const numId = isBullet
            ? numberingManager.createBulletList()
            : numberingManager.createNumberedList();
          correctedNumIdByLevel.set(key, numId);

          const instance = numberingManager.getInstance(numId);
          if (instance) {
            const abstractNum = numberingManager.getAbstractNumbering(instance.getAbstractNumId());
            if (abstractNum) {
              applyIndentationSettings(abstractNum, options.indentationLevels!, isBullet);
            }
          }
        }
        return correctedNumIdByLevel.get(key)!;
      };

      for (const item of analysis.paragraphs) {
        if (item.detection.isWordList && item.detection.numId !== null) {
          const para = item.paragraph as Paragraph;
          const numbering = para.getNumbering();
          if (numbering) {
            const isBullet = item.detection.category === "bullet";
            const correctedNumId = getCorrectedNumId(numbering.level, isBullet);
            para.setNumbering(correctedNumId, numbering.level);
            report.normalized++;
            report.details.push({
              originalText: item.text.substring(0, 50),
              action: "normalized",
              reason: `Applied indentation settings at level ${numbering.level}`,
            });
          }
        }
      }
      normalizeOrphanListLevelsInCell(cell);
      return report;
    }
    report.skipped = analysis.paragraphs.length;
    // Always normalize orphan levels even when no other normalization needed
    normalizeOrphanListLevelsInCell(cell);
    return report;
  }

  // Calculate baseline (minimum) indentation for relative level inference
  let baselineIndent = Infinity;
  for (const item of analysis.paragraphs) {
    if (item.detection.category !== "none") {
      baselineIndent = Math.min(baselineIndent, item.detection.indentationTwips);
    }
  }
  if (baselineIndent === Infinity) baselineIndent = 0;

  // Calculate level shifts PER LIST GROUP based on MAJORITY CATEGORY items only
  // This ensures minority category items don't affect the shift calculation
  // A "list group" is a contiguous sequence of list items separated by non-list items
  const levelShiftByIndex = new Map<number, number>();
  let currentGroupStart = -1;
  let currentGroupMinLevel = Infinity;

  for (let i = 0; i < analysis.paragraphs.length; i++) {
    const item = analysis.paragraphs[i]!;

    // Only consider majority category items for level shift calculation
    if (item.detection.category === majorityCategory) {
      if (currentGroupStart === -1) {
        currentGroupStart = i; // Start new group
        currentGroupMinLevel = Infinity;
      }
      // Track minimum level in current group (only from majority category)
      currentGroupMinLevel = Math.min(currentGroupMinLevel, item.detection.inferredLevel);
    } else if (item.detection.category === "none") {
      // Non-list item - end current group if any
      if (currentGroupStart !== -1) {
        // Apply the group's level shift to ALL non-"none" items in the group
        const shift = currentGroupMinLevel === Infinity ? 0 : currentGroupMinLevel;
        for (let j = currentGroupStart; j < i; j++) {
          if (analysis.paragraphs[j]!.detection.category !== "none") {
            levelShiftByIndex.set(j, shift);
          }
        }
        currentGroupStart = -1;
        currentGroupMinLevel = Infinity;
      }
    }
    // Minority category items don't break the group but don't affect minLevel
  }

  // Handle last group if cell ends with list items
  if (currentGroupStart !== -1) {
    const shift = currentGroupMinLevel === Infinity ? 0 : currentGroupMinLevel;
    for (let j = currentGroupStart; j < analysis.paragraphs.length; j++) {
      if (analysis.paragraphs[j]!.detection.category !== "none") {
        levelShiftByIndex.set(j, shift);
      }
    }
  }

  // === Context-aware sub-item detection ===
  // Track which items should be treated as sub-items and their parent indices
  const bulletAsSubItemIndices = new Set<number>();
  const numberedAsSubItemIndices = new Set<number>();
  const parentIndexByIndex = new Map<number, number>();

  // Helper to calculate normalized level for an item (used for parent level lookup)
  const getNormalizedLevel = (itemIndex: number): number => {
    const item = analysis.paragraphs[itemIndex]!;
    const detection = item.detection;
    const hasTypedPrefix = !!detection.typedPrefix;
    const levelShift = levelShiftByIndex.get(itemIndex) ?? 0;

    if (hasTypedPrefix) {
      const relativeIndent = detection.indentationTwips - baselineIndent;
      const rawLevel = inferLevelFromRelativeIndentation(relativeIndent);
      // Apply levelShift consistently for typed prefixes too
      return Math.max(0, rawLevel - levelShift);
    } else {
      return Math.max(0, detection.inferredLevel - levelShift);
    }
  };

  // Minimum indentation difference (in twips) to consider an item a sub-item
  // 200 twips ≈ 0.14 inches - small enough to catch real sub-items but avoid false positives
  const INDENT_THRESHOLD = 200;

  if (majorityCategory === "numbered") {
    let lastNumberedItemIndex = -1;

    for (let i = 0; i < analysis.paragraphs.length; i++) {
      const item = analysis.paragraphs[i]!;
      const detection = item.detection;

      if (detection.category === "numbered") {
        lastNumberedItemIndex = i;
      } else if (detection.category === "bullet" && lastNumberedItemIndex >= 0) {
        // Only mark as sub-item if actually indented MORE than the parent
        // This prevents level-0 bullets from being wrongly demoted to level-1
        const parentDetection = analysis.paragraphs[lastNumberedItemIndex]!.detection;
        if (detection.indentationTwips > parentDetection.indentationTwips + INDENT_THRESHOLD) {
          bulletAsSubItemIndices.add(i);
          parentIndexByIndex.set(i, lastNumberedItemIndex);
        } else if (detection.indentationTwips === 0 && parentDetection.indentationTwips === 0) {
          // Fallback: when both have 0 indentation (no explicit w:ind), treat minority
          // category items as sub-items. This handles documents where indentation comes
          // from numbering definitions rather than explicit paragraph indentation.
          bulletAsSubItemIndices.add(i);
          parentIndexByIndex.set(i, lastNumberedItemIndex);
        }
        // If not indented more, treat as separate list (don't add to sub-item set)
      } else if (detection.category === "none") {
        lastNumberedItemIndex = -1;
      }
    }
  }

  if (majorityCategory === "bullet") {
    let lastBulletItemIndex = -1;

    for (let i = 0; i < analysis.paragraphs.length; i++) {
      const item = analysis.paragraphs[i]!;
      const detection = item.detection;

      if (detection.category === "bullet") {
        lastBulletItemIndex = i;
      } else if (detection.category === "numbered" && lastBulletItemIndex >= 0) {
        // Only mark as sub-item if actually indented MORE than the parent
        const parentDetection = analysis.paragraphs[lastBulletItemIndex]!.detection;
        if (detection.indentationTwips > parentDetection.indentationTwips + INDENT_THRESHOLD) {
          numberedAsSubItemIndices.add(i);
          parentIndexByIndex.set(i, lastBulletItemIndex);
        } else if (detection.indentationTwips === 0 && parentDetection.indentationTwips === 0) {
          // Fallback: when both have 0 indentation (no explicit w:ind), treat minority
          // category items as sub-items. This handles documents where indentation comes
          // from numbering definitions rather than explicit paragraph indentation.
          numberedAsSubItemIndices.add(i);
          parentIndexByIndex.set(i, lastBulletItemIndex);
        }
      } else if (detection.category === "none") {
        lastBulletItemIndex = -1;
      }
    }
  }
  // === End sub-item detection ===

  // Track numId per level - will be reset when parent level appears
  const numIdByLevel = new Map<number, number>();
  let lastProcessedLevel = -1;

  // Helper to get/create numId for a level (uses majority category)
  const getNumId = (level: number): number => {
    if (level < lastProcessedLevel) {
      for (const existingLevel of numIdByLevel.keys()) {
        if (existingLevel > level) {
          numIdByLevel.delete(existingLevel);
        }
      }
    }
    lastProcessedLevel = level;

    if (!numIdByLevel.has(level)) {
      const numId =
        majorityCategory === "numbered"
          ? numberingManager.createNumberedList()
          : numberingManager.createBulletList();
      numIdByLevel.set(level, numId);

      // Apply user's indentation settings if provided
      if (options?.indentationLevels?.length) {
        const instance = numberingManager.getInstance(numId);
        if (instance) {
          const abstractNum = numberingManager.getAbstractNumbering(instance.getAbstractNumId());
          if (abstractNum) {
            applyIndentationSettings(abstractNum, options.indentationLevels, majorityCategory !== "numbered");
          }
        }
      }
    }
    return numIdByLevel.get(level)!;
  };

  // Separate tracking for bullet numIds (used for trailing bullets in numbered-majority cells)
  const bulletNumIdByLevel = new Map<number, number>();
  let lastBulletProcessedLevel = -1;

  const getBulletNumId = (level: number): number => {
    if (level < lastBulletProcessedLevel) {
      for (const existingLevel of bulletNumIdByLevel.keys()) {
        if (existingLevel > level) {
          bulletNumIdByLevel.delete(existingLevel);
        }
      }
    }
    lastBulletProcessedLevel = level;

    if (!bulletNumIdByLevel.has(level)) {
      const numId = numberingManager.createBulletList();
      bulletNumIdByLevel.set(level, numId);

      // Apply user's indentation settings if provided
      if (options?.indentationLevels?.length) {
        const instance = numberingManager.getInstance(numId);
        if (instance) {
          const abstractNum = numberingManager.getAbstractNumbering(instance.getAbstractNumId());
          if (abstractNum) {
            applyIndentationSettings(abstractNum, options.indentationLevels, true);
          }
        }
      }
    }
    return bulletNumIdByLevel.get(level)!;
  };

  // Process each paragraph
  for (let index = 0; index < analysis.paragraphs.length; index++) {
    const item = analysis.paragraphs[index]!;
    const { paragraph, text, detection } = item;
    const para = paragraph as Paragraph;

    // Skip non-list items entirely - preserve "Note:", plain text, etc.
    if (detection.category === "none") {
      report.skipped++;
      report.details.push({
        originalText: text.substring(0, 50),
        action: "skipped",
        reason: "Not a list item - preserving original formatting",
      });
      continue;
    }

    try {
      // Check if this item needs conversion (different category than majority)
      const needsConversion = detection.category !== majorityCategory;
      const hasTypedPrefix = !!detection.typedPrefix;
      const isWordList = detection.isWordList;

      // Get the level shift for this paragraph's list group
      const levelShift = levelShiftByIndex.get(index) ?? 0;

      // Calculate target level
      // - For typed prefixes: use format-based level (decimal=0, letter=1, roman=2) with shift
      //   unless explicitly indented, in which case use indentation-based level
      // - For sub-items: use parent's normalized level + 1
      // - For Word lists: use format-based level with level shift applied
      let targetLevel: number;
      if (hasTypedPrefix) {
        const relativeIndent = detection.indentationTwips - baselineIndent;
        const indentBasedLevel = inferLevelFromRelativeIndentation(relativeIndent);
        // Use format-based level (from FORMAT_TO_LEVEL) if no explicit extra indentation
        // This preserves semantic hierarchy: 1.→L0, a)→L1, i.→L2
        if (indentBasedLevel === 0 && detection.inferredLevel > 0) {
          targetLevel = Math.max(0, detection.inferredLevel - levelShift);
        } else {
          targetLevel = indentBasedLevel;
        }
      } else if (bulletAsSubItemIndices.has(index) || numberedAsSubItemIndices.has(index)) {
        // Sub-item: use parent's NORMALIZED level + 1
        const parentIndex = parentIndexByIndex.get(index);
        const parentNormalizedLevel = parentIndex !== undefined ? getNormalizedLevel(parentIndex) : 0;
        targetLevel = parentNormalizedLevel + 1;
      } else {
        targetLevel = Math.max(0, detection.inferredLevel - levelShift);
      }

      // Process based on what type of item this is
      if (hasTypedPrefix && detection.typedPrefix) {
        // Typed prefix: strip prefix and apply new formatting
        stripTypedPrefix(para, detection.typedPrefix);
        para.setNumbering(getNumId(targetLevel), targetLevel);
        report.normalized++;
        report.details.push({
          originalText: text.substring(0, 50),
          action: "normalized",
          reason: `Typed prefix → level ${targetLevel}`,
        });
      } else if (isWordList && bulletAsSubItemIndices.has(index)) {
        // Sandwiched bullet following numbered → convert to numbered sub-item
        para.setNumbering(getNumId(targetLevel), targetLevel);
        report.normalized++;
        report.details.push({
          originalText: text.substring(0, 50),
          action: "normalized",
          reason: `Bullet → numbered sub-item at level ${targetLevel}`,
        });
      } else if (isWordList && numberedAsSubItemIndices.has(index)) {
        // Numbered following bullet → convert to bullet sub-item
        para.setNumbering(getBulletNumId(targetLevel), targetLevel);
        report.normalized++;
        report.details.push({
          originalText: text.substring(0, 50),
          action: "normalized",
          reason: `Numbered → bullet at level ${targetLevel}`,
        });
      } else if (isWordList && detection.category === "bullet" && majorityCategory === "numbered" && !bulletAsSubItemIndices.has(index)) {
        // Trailing bullet in numbered-majority cell - preserve as bullet
        para.setNumbering(getBulletNumId(targetLevel), targetLevel);
        report.normalized++;
        report.details.push({
          originalText: text.substring(0, 50),
          action: "normalized",
          reason: `Trailing bullet preserved at level ${targetLevel}`,
        });
      } else if (isWordList && needsConversion) {
        // Regular category conversion
        if (majorityCategory === "bullet") {
          para.setNumbering(getBulletNumId(targetLevel), targetLevel);
        } else {
          para.setNumbering(getNumId(targetLevel), targetLevel);
        }
        report.normalized++;
        report.details.push({
          originalText: text.substring(0, 50),
          action: "normalized",
          reason: `Word ${detection.category} → ${majorityCategory} level ${targetLevel}`,
        });
      } else if (isWordList) {
        // Preserve category but ensure consistent numId with user settings
        if (detection.category === "bullet") {
          para.setNumbering(getBulletNumId(targetLevel), targetLevel);
        } else {
          para.setNumbering(getNumId(targetLevel), targetLevel);
        }
        report.normalized++;
        report.details.push({
          originalText: text.substring(0, 50),
          action: "normalized",
          reason: `Updated numId for consistent numbering at level ${targetLevel}`,
        });
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      report.errors.push(`Failed on "${text.substring(0, 30)}...": ${message}`);
      report.details.push({
        originalText: text.substring(0, 50),
        action: "error",
        reason: message,
      });
    }
  }

  // Ensure list items don't start at orphan levels (level 1+ without level 0 parent)
  // This handles edge cases where the level shift calculation still results in non-zero levels
  normalizeOrphanListLevelsInCell(cell);

  return report;
}

/**
 * Normalize lists across all cells in a table.
 */
export function normalizeListsInTable(
  table: Table,
  options: ListNormalizationOptions,
  numberingManager: NumberingManager
): ListNormalizationReport {
  const aggregateReport: ListNormalizationReport = {
    normalized: 0,
    skipped: 0,
    errors: [],
    appliedCategory: "none",
    details: [],
  };

  for (const row of table.getRows()) {
    for (const cell of row.getCells()) {
      const cellReport = normalizeListsInCell(cell, options, numberingManager);

      aggregateReport.normalized += cellReport.normalized;
      aggregateReport.skipped += cellReport.skipped;
      aggregateReport.errors.push(...cellReport.errors);
      aggregateReport.details.push(...cellReport.details);

      if (cellReport.appliedCategory !== "none") {
        aggregateReport.appliedCategory = cellReport.appliedCategory;
      }
    }
  }

  return aggregateReport;
}

/**
 * Normalize orphan Level 1+ list items in a table cell.
 *
 * Detects when a cell's first list item starts at Level 1 or higher (e.g., open circles)
 * without a preceding Level 0 item (e.g., filled circles). This is common when content
 * is extracted from a larger document where the Level 0 parent existed elsewhere.
 *
 * The function shifts all list items down by the minimum level found, so they start at Level 0.
 *
 * @param cell - The table cell to normalize
 * @returns Number of paragraphs that were adjusted
 *
 * @example
 * // Before: Cell has Level 1 bullets (○ Name, ○ Address, ○ Phone)
 * const count = normalizeOrphanListLevelsInCell(cell);
 * // After: Cell has Level 0 bullets (● Name, ● Address, ● Phone)
 */
export function normalizeOrphanListLevelsInCell(cell: TableCell): number {
  const paragraphs = cell.getParagraphs();

  // Find minimum level among all list items in the cell
  let minLevel = Infinity;
  let hasListItems = false;

  for (const para of paragraphs) {
    const numbering = para.getNumbering();
    if (numbering) {
      hasListItems = true;
      minLevel = Math.min(minLevel, numbering.level);
    }
  }

  // If no list items or already at Level 0, nothing to fix
  if (!hasListItems || minLevel === 0 || minLevel === Infinity) {
    return 0;
  }

  // Shift all list items down by minLevel
  let normalizedCount = 0;
  for (const para of paragraphs) {
    const numbering = para.getNumbering();
    if (numbering) {
      const newLevel = numbering.level - minLevel;
      para.setNumbering(numbering.numId, newLevel);
      normalizedCount++;
    }
  }

  return normalizedCount;
}

/**
 * Normalize orphan Level 1+ list items across all cells in a table.
 *
 * @param table - The table to normalize
 * @returns Total number of paragraphs adjusted across all cells
 */
export function normalizeOrphanListLevelsInTable(table: Table): number {
  let totalNormalized = 0;

  for (const row of table.getRows()) {
    for (const cell of row.getCells()) {
      totalNormalized += normalizeOrphanListLevelsInCell(cell);
    }
  }

  return totalNormalized;
}

// =============================================================================
// NUMBERING MANAGER HELPERS
// =============================================================================

/**
 * Get existing or create new numbered list numId.
 */
function getOrCreateNumberedListNumId(
  numberingManager: NumberingManager
): number {
  // First, try to find an existing numbered list
  const instances = numberingManager.getAllInstances();
  for (const instance of instances) {
    const abstractNum = numberingManager.getAbstractNumbering(
      instance.getAbstractNumId()
    );
    if (abstractNum) {
      const level0 = abstractNum.getLevel(0);
      if (level0) {
        const format = level0.getFormat();
        if (getListCategoryFromFormat(format) === "numbered") {
          return instance.getNumId();
        }
      }
    }
  }

  // Create a new numbered list
  return numberingManager.createNumberedList();
}

/**
 * Get existing or create new bullet list numId.
 */
function getOrCreateBulletListNumId(
  numberingManager: NumberingManager
): number {
  // First, try to find an existing bullet list
  const instances = numberingManager.getAllInstances();
  for (const instance of instances) {
    const abstractNum = numberingManager.getAbstractNumbering(
      instance.getAbstractNumId()
    );
    if (abstractNum) {
      const level0 = abstractNum.getLevel(0);
      if (level0) {
        const format = level0.getFormat();
        if (format === "bullet") {
          return instance.getNumId();
        }
      }
    }
  }

  // Create a new bullet list
  return numberingManager.createBulletList();
}

// =============================================================================
// PUBLIC API CLASS
// =============================================================================

/**
 * Main entry point for list normalization.
 *
 * @example
 * ```typescript
 * const normalizer = new ListNormalizer(numberingManager);
 *
 * // Analyze a cell
 * const analysis = normalizer.analyzeCell(cellElement);
 * console.log(`Has typed lists: ${analysis.hasTypedLists}`);
 * console.log(`Majority type: ${analysis.majorityCategory}`);
 *
 * // Normalize a cell
 * const report = normalizer.normalizeCell(cellElement, {
 *   numberedStyleNumId: 5,
 *   bulletStyleNumId: 8,
 *   scope: 'cell',
 * });
 * console.log(`Normalized ${report.normalized} items`);
 * ```
 */
export class ListNormalizer {
  private numberingManager: NumberingManager;

  constructor(numberingManager: NumberingManager) {
    this.numberingManager = numberingManager;
  }

  /**
   * Analyze lists in a cell.
   */
  analyzeCell(cell: TableCell): ListAnalysis {
    return analyzeCellLists(cell);
  }

  /**
   * Analyze lists in a table.
   */
  analyzeTable(table: Table): Map<TableCell, ListAnalysis> {
    return analyzeTableLists(table);
  }

  /**
   * Normalize lists in a cell.
   */
  normalizeCell(
    cell: TableCell,
    options: Partial<ListNormalizationOptions> = {}
  ): ListNormalizationReport {
    const fullOptions = this.resolveOptions(options);
    return normalizeListsInCell(cell, fullOptions, this.numberingManager);
  }

  /**
   * Normalize lists in a table.
   */
  normalizeTable(
    table: Table,
    options: Partial<ListNormalizationOptions> = {}
  ): ListNormalizationReport {
    const fullOptions = this.resolveOptions(options);
    return normalizeListsInTable(table, fullOptions, this.numberingManager);
  }

  /**
   * Normalize lists in all tables.
   */
  normalizeAllTables(
    tables: Table[],
    options: Partial<ListNormalizationOptions> = {}
  ): ListNormalizationReport {
    const aggregateReport: ListNormalizationReport = {
      normalized: 0,
      skipped: 0,
      errors: [],
      appliedCategory: "none",
      details: [],
    };

    for (const table of tables) {
      const tableReport = this.normalizeTable(table, options);
      aggregateReport.normalized += tableReport.normalized;
      aggregateReport.skipped += tableReport.skipped;
      aggregateReport.errors.push(...tableReport.errors);
      aggregateReport.details.push(...tableReport.details);

      if (tableReport.appliedCategory !== "none") {
        aggregateReport.appliedCategory = tableReport.appliedCategory;
      }
    }

    if (aggregateReport.normalized > 0) {
      defaultLogger.info(
        `List normalization complete: ${aggregateReport.normalized} items normalized`
      );
    }

    return aggregateReport;
  }

  /**
   * Resolve partial options with defaults.
   */
  private resolveOptions(
    partial: Partial<ListNormalizationOptions>
  ): ListNormalizationOptions {
    return {
      numberedStyleNumId:
        partial.numberedStyleNumId ??
        getOrCreateNumberedListNumId(this.numberingManager),
      bulletStyleNumId:
        partial.bulletStyleNumId ??
        getOrCreateBulletListNumId(this.numberingManager),
      scope: partial.scope ?? "cell",
      forceMajority: partial.forceMajority ?? false,
      preserveIndentation: partial.preserveIndentation ?? false,
      indentationLevels: partial.indentationLevels,
    };
  }
}
