# Style Parameter Enhancement for ensureBlankLines\* Functions

## Summary

Successfully exposed the `style` parameter for both [`ensureBlankLinesAfter1x1Tables()`](src/core/Document.ts:3775) and [`ensureBlankLinesAfterOtherTables()`](src/core/Document.ts:3903) functions, allowing users to configure the paragraph style for blank lines added after tables.

## Changes Made

### 1. Updated Function Signatures

Both functions now accept an optional `style` parameter in their options:

```typescript
ensureBlankLinesAfter1x1Tables(options?: {
  spacingAfter?: number;
  markAsPreserved?: boolean;
  style?: string;  // NEW
  filter?: (table: Table, index: number) => boolean;
})

ensureBlankLinesAfterOtherTables(options?: {
  spacingAfter?: number;
  markAsPreserved?: boolean;
  style?: string;  // NEW
  filter?: (table: Table, index: number) => boolean;
})
```

### 2. Default Behavior

- **Default style**: `'Normal'` (maintains backward compatibility)
- All existing code continues to work without modification
- Style parameter is completely optional

### 3. Implementation Details

Updated both functions to:

- Extract the `style` parameter from options with default of `'Normal'`
- Apply the specified style to all newly created blank paragraphs
- Use `blankPara.setStyle(style)` instead of hardcoded `blankPara.setStyle("Normal")`

#### Lines Changed in [`src/core/Document.ts`](src/core/Document.ts):

- Line 3738: Added `style` parameter to options interface
- Line 3787: Extract style with default value
- Line 3833, 3844: Apply configurable style to blank paragraphs
- Line 3873: Added `style` parameter to options interface
- Line 3915: Extract style with default value
- Line 3961, 3972: Apply configurable style to blank paragraphs

### 4. Documentation Updates

Enhanced JSDoc documentation for both functions with new examples:

```typescript
// Example: Custom style for blank paragraphs
doc.ensureBlankLinesAfter1x1Tables({
  style: "BodyText", // Use BodyText instead of Normal
  spacingAfter: 120,
});

// Example: Different styles for different table types
doc.ensureBlankLinesAfter1x1Tables({ style: "Normal" });
doc.ensureBlankLinesAfterOtherTables({ style: "BodyText" });
```

### 5. Example File Updates

Updated [`examples/ensure-blank-lines-after-tables.ts`](examples/ensure-blank-lines-after-tables.ts) with:

- Example 1: Basic usage with explicit style parameter
- Example 2: Header 2 tables with custom style
- Example 3: Multi-cell tables with custom style
- Example 4: Different styles for different table types

### 6. Test Coverage

Created [`src/__tests__/ensure-blank-lines-style.test.ts`](src/__tests__/ensure-blank-lines-style.test.ts) with comprehensive tests:

- Default style behavior (Normal)
- Custom style application
- Combined style and spacing configuration
- 1x1 vs multi-cell table distinction
- Backward compatibility verification

## Backward Compatibility ✓

**FULLY BACKWARD COMPATIBLE**

- Style parameter is optional with sensible default (`'Normal'`)
- All existing code continues to work without changes
- No breaking changes to API surface
- Behavior only changes when explicitly configured

## Usage Examples

### Basic Usage (Unchanged)

```typescript
// Works exactly as before - uses Normal style
const result = doc.ensureBlankLinesAfter1x1Tables();
```

### Custom Style

```typescript
// Use custom style for blank lines
const result = doc.ensureBlankLinesAfter1x1Tables({
  style: "BodyText",
  spacingAfter: 120,
});
```

### Different Styles Per Table Type

```typescript
// Normal for 1x1 tables (Header 2)
doc.ensureBlankLinesAfter1x1Tables({
  style: "Normal",
  spacingAfter: 120,
});

// BodyText for multi-cell tables
doc.ensureBlankLinesAfterOtherTables({
  style: "BodyText",
  spacingAfter: 180,
});
```

## Benefits

1. **Flexibility**: Users can now match blank line styles to their document's style system
2. **Consistency**: Allows maintaining consistent styling conventions across document types
3. **No Breaking Changes**: All existing code continues to work
4. **Clear API**: Simple, intuitive parameter naming
5. **Well-Documented**: Comprehensive examples in code and documentation

## Files Modified

- [`src/core/Document.ts`](src/core/Document.ts) - Core implementation
- [`examples/ensure-blank-lines-after-tables.ts`](examples/ensure-blank-lines-after-tables.ts) - Usage examples
- [`src/__tests__/ensure-blank-lines-style.test.ts`](src/__tests__/ensure-blank-lines-style.test.ts) - Test coverage

## Note on Prettier

Prettier is auto-formatting code according to project configuration (converting double quotes to single quotes). This is expected behavior and doesn't affect functionality - it's just code style enforcement.
