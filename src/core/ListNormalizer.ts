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
} from "../types/list-types";
import {
  detectListType,
  detectTypedPrefix,
  getListCategoryFromFormat,
} from "../utils/list-detection";
import { defaultLogger } from "../utils/logger";
import { isRun } from "../elements/Paragraph";

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
 */
export function normalizeListsInCell(
  cell: TableCell,
  options: Required<ListNormalizationOptions>,
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

  // Nothing to do if no normalization needed
  if (analysis.recommendedAction === "none") {
    report.skipped = analysis.paragraphs.length;
    return report;
  }

  // Track numId per level to ensure proper sequencing
  const numIdByLevel = new Map<number, number>();

  // Helper to get/create numId for a level
  const getNumId = (level: number): number => {
    if (!numIdByLevel.has(level)) {
      const numId =
        majorityCategory === "numbered"
          ? numberingManager.createNumberedList()
          : numberingManager.createBulletList();
      numIdByLevel.set(level, numId);
    }
    return numIdByLevel.get(level)!;
  };

  // Process each paragraph
  for (const item of analysis.paragraphs) {
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

      // Calculate target level
      // - Use format-based level from detection (decimal=0, letter=1, roman=2)
      // - Word bullets being converted get level 1
      let targetLevel = detection.inferredLevel;
      if (needsConversion && isWordList) {
        targetLevel = 1; // Word bullets → numbered level 1 (a., b., c.)
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
      } else if (isWordList && needsConversion) {
        // Word list that doesn't match majority: convert it
        para.setNumbering(getNumId(targetLevel), targetLevel);
        report.normalized++;
        report.details.push({
          originalText: text.substring(0, 50),
          action: "normalized",
          reason: `Word ${detection.category} → ${majorityCategory} level ${targetLevel}`,
        });
      } else if (isWordList) {
        // Word list already matches majority - skip
        report.skipped++;
        report.details.push({
          originalText: text.substring(0, 50),
          action: "skipped",
          reason: "Word list already matches majority type",
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

  return report;
}

/**
 * Normalize lists across all cells in a table.
 */
export function normalizeListsInTable(
  table: Table,
  options: Required<ListNormalizationOptions>,
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
  ): Required<ListNormalizationOptions> {
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
    };
  }
}
