/**
 * Helper functions for creating complex and nested fields
 */

import { ComplexField } from './Field';

/**
 * Creates a nested IF field containing a MERGEFIELD
 * This is a common pattern for conditional mail merge
 *
 * @param condition The IF condition (e.g., 'Status = "Active"')
 * @param mergeFieldName The merge field name to include
 * @param trueText Text to show when condition is true
 * @param falseText Text to show when condition is false
 * @returns ComplexField with nested MERGEFIELD
 *
 * @example
 * ```typescript
 * const field = createNestedIFMergeField('Status = "Active"', 'Name', 'Active: ', 'Inactive');
 * // Result: IF field that shows "Active: [Name]" if Status is Active, otherwise "Inactive"
 * ```
 */
export function createNestedIFMergeField(
  condition: string,
  mergeFieldName: string,
  trueText: string = '',
  falseText: string = ''
): ComplexField {
  // Create the nested MERGEFIELD
  const mergeField = new ComplexField({
    instruction: ` MERGEFIELD ${mergeFieldName} `,
    result: `[${mergeFieldName}]`,
  });

  // Create the IF field with nested MERGEFIELD
  const ifField = new ComplexField({
    instruction: ` IF ${condition} "${trueText}" "${falseText}" `,
    result: trueText || falseText,
  });

  ifField.addNestedField(mergeField);

  return ifField;
}

/**
 * Creates a MERGEFIELD with custom formatting
 *
 * @param fieldName The merge field name
 * @param format Optional format switches
 * @returns ComplexField for MERGEFIELD
 *
 * @example
 * ```typescript
 * const field = createMergeField('Date', '\\@ "MMMM d, yyyy"');
 * ```
 */
export function createMergeField(fieldName: string, format?: string): ComplexField {
  let instruction = ` MERGEFIELD ${fieldName}`;

  if (format) {
    instruction += ` ${format}`;
  }

  instruction += ' \\* MERGEFORMAT ';

  return new ComplexField({
    instruction,
    result: `[${fieldName}]`,
  });
}

/**
 * Creates a REF field with a nested field in the bookmark reference
 * Used for complex cross-references
 *
 * @param bookmarkName The bookmark to reference
 * @param format Optional format switches
 * @returns ComplexField for REF
 *
 * @example
 * ```typescript
 * const field = createRefField('Chapter1', '\\h');
 * ```
 */
export function createRefField(bookmarkName: string, format?: string): ComplexField {
  let instruction = ` REF ${bookmarkName}`;

  if (format) {
    instruction += ` ${format}`;
  } else {
    instruction += ' \\h'; // Hyperlink by default
  }

  instruction += ' \\* MERGEFORMAT ';

  return new ComplexField({
    instruction,
    result: `[${bookmarkName}]`,
  });
}

/**
 * Creates an IF field with custom true/false branches
 *
 * @param condition The condition to evaluate
 * @param trueContent Content to show when true
 * @param falseContent Content to show when false
 * @returns ComplexField for IF
 *
 * @example
 * ```typescript
 * const field = createIFField('Amount > 1000', 'High Value', 'Normal');
 * ```
 */
export function createIFField(
  condition: string,
  trueContent: string,
  falseContent: string = ''
): ComplexField {
  const instruction = ` IF ${condition} "${trueContent}" "${falseContent}" `;

  return new ComplexField({
    instruction,
    result: trueContent,
  });
}

/**
 * Creates a complex nested field structure with multiple levels
 * Useful for advanced scenarios like nested IF statements
 *
 * @param outerInstruction The outer field instruction
 * @param nestedFields Array of nested fields to include
 * @returns ComplexField with nested structure
 *
 * @example
 * ```typescript
 * const innerField = createMergeField('Amount');
 * const field = createNestedField('IF Amount > 0', [innerField]);
 * ```
 */
export function createNestedField(
  outerInstruction: string,
  nestedFields: ComplexField[]
): ComplexField {
  const field = new ComplexField({
    instruction: outerInstruction,
  });

  for (const nested of nestedFields) {
    field.addNestedField(nested);
  }

  return field;
}
