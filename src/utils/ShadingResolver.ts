/**
 * ShadingResolver - Resolves effective cell shading through the ECMA-376 style hierarchy
 *
 * Per ECMA-376 §17.7.6, cell shading is resolved in priority order:
 * 1. Direct cell shading
 * 2. Row table property exceptions (tblPrEx)
 * 3. Conditional table style (tblStylePr) based on cnfStyle bitmask
 * 4. Table style default cell shading
 * 5. Direct table shading
 *
 * Special values:
 * - pattern: "nil" = explicitly clear (stop resolution)
 * - fill: "auto" = inherit (continue resolution)
 */

import type { ShadingConfig } from '../elements/CommonTypes';
import type { Table } from '../elements/Table';
import type { TableCell } from '../elements/TableCell';
import type { StylesManager } from '../formatting/StylesManager';
import { getActiveConditionalsInPriorityOrder } from './cnfStyleDecoder';
import type { ConditionalFormattingType } from '../formatting/Style';

/**
 * Resolves the effective shading for a table cell through the style hierarchy.
 *
 * @param cell - The table cell to resolve shading for
 * @param table - The parent table
 * @param stylesManager - The document's styles manager for style lookup
 * @returns The resolved ShadingConfig, or undefined if no shading applies
 */
export function resolveCellShading(
  cell: TableCell,
  table: Table,
  stylesManager: StylesManager
): ShadingConfig | undefined {
  // 1. Direct cell shading
  const directShading = cell.getShading();
  if (directShading) {
    if (directShading.pattern === 'nil') return undefined;
    if (directShading.fill && directShading.fill !== 'auto') return directShading;
  }

  // 2. Row table property exceptions
  const row = findRowForCell(cell, table);
  if (row) {
    const exceptions = row.getTablePropertyExceptions();
    if (exceptions?.shading) {
      if (exceptions.shading.pattern === 'nil') return undefined;
      if (exceptions.shading.fill && exceptions.shading.fill !== 'auto') return exceptions.shading;
    }
  }

  // 3. Conditional table style shading
  const tableStyleId = table.getStyle();
  if (tableStyleId) {
    const style = stylesManager.getStyle(tableStyleId);
    if (style) {
      const styleProps = style.getProperties();
      const conditionals = styleProps.tableStyle?.conditionalFormatting;
      if (conditionals) {
        const cnfStyle = cell.getCnfStyle();
        if (cnfStyle) {
          const activeConditionals = getActiveConditionalsInPriorityOrder(cnfStyle);
          for (const condType of activeConditionals) {
            const match = conditionals.find((c) => c.type === condType);
            if (match?.cellFormatting?.shading) {
              const condShading = match.cellFormatting.shading;
              if (condShading.pattern === 'nil') return undefined;
              return condShading;
            }
          }
        }
      }

      // 4. Table style default cell shading
      if (styleProps.tableStyle?.cell?.shading) {
        const defaultCellShading = styleProps.tableStyle.cell.shading;
        if (defaultCellShading.pattern === 'nil') return undefined;
        return defaultCellShading;
      }
    }
  }

  // 5. Direct table shading
  const tableShading = table.getShading();
  if (tableShading) {
    if (tableShading.pattern === 'nil') return undefined;
    return tableShading;
  }

  return undefined;
}

/**
 * Finds the row that contains a given cell.
 */
function findRowForCell(cell: TableCell, table: Table) {
  for (const row of table.getRows()) {
    const cells = row.getCells();
    if (cells.includes(cell)) {
      return row;
    }
  }
  return undefined;
}
