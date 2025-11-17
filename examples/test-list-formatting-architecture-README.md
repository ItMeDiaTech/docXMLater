# List Formatting Architecture Demo

This example demonstrates the **3-phase list formatting architecture** using real docXMLater APIs.

## Architecture Overview

The list formatting system follows a clear separation of concerns across 3 phases:

### Phase 1: Normalize List Types

Standardizes the structure and format of lists:

- [`normalizeBulletLists()`](../src/core/Document.ts:2487) - Converts all bullet lists to standard hierarchy: • ○ ■
- [`normalizeNumberedLists()`](../src/core/Document.ts:2421) - Converts all numbered lists to standard formats: 1. a. i.

### Phase 2: Standardize Formatting

Applies consistent visual formatting to list symbols/prefixes:

- [`standardizeBulletSymbols()`](../src/core/Document.ts:4528) - Formats bullet symbols (font, size, color, bold)
- [`standardizeNumberedListPrefixes()`](../src/core/Document.ts:4607) - Formats number prefixes (font, size, color, bold)

### Phase 3: Normalize Indentation

Ensures consistent spacing and indentation:

- [`normalizeAllListIndentation()`](../src/core/Document.ts:2834) - Applies standard indentation (0.5" increments, 0.25" hanging)

## Test Scenarios

### Scenario 1: Complete Normalization

Applies all three phases to fully standardize list formatting.

- **Output**: `test-list-formatted-scenario1.docx`
- **Use Case**: Complete document cleanup and standardization

### Scenario 2: Custom Red Bullets

Normalizes bullet lists only with custom red formatting.

- **Output**: `test-list-formatted-scenario2.docx`
- **Use Case**: Brand-specific bullet styling (e.g., company red)
- **Formatting**: Calibri 14pt bold red

### Scenario 3: Professional Blue Numbers

Normalizes numbered lists only with professional blue formatting.

- **Output**: `test-list-formatted-scenario3.docx`
- **Use Case**: Professional documents (legal, academic)
- **Formatting**: Times New Roman 12pt blue

### Scenario 4: Indentation Only

Normalizes only indentation while preserving all existing styles.

- **Output**: `test-list-formatted-scenario4.docx`
- **Use Case**: Fix spacing issues without changing visual appearance
- **Preserves**: Bullet characters, numbering formats, fonts, colors

## Usage

### Run All Scenarios

```bash
npm run test-list-arch
# or
npm run test-list-arch all
```

### Run Specific Scenario

```bash
npm run test-list-arch scenario 1   # Complete normalization
npm run test-list-arch scenario 2   # Custom red bullets
npm run test-list-arch scenario 3   # Professional blue numbers
npm run test-list-arch scenario 4   # Indentation only
```

### Run Directly with ts-node

```bash
npx ts-node examples/test-list-formatting-architecture.ts
npx ts-node examples/test-list-formatting-architecture.ts scenario 2
```

## Expected Output

Each scenario will:

1. Load `Test_Code.docx` from the project root
2. Display a detailed execution plan showing which phases will run
3. Execute the selected scenario with real-time logging
4. Save the result to `examples/output/test-list-formatted-scenarioN.docx`
5. Display statistics about the changes made

### Sample Console Output

```
================================================================================
LIST FORMATTING ARCHITECTURE - EXECUTION PLAN
================================================================================

PHASE 1: NORMALIZE LIST TYPES
  ├─ normalizeBulletLists()      → Standardize all bullet lists to • ○ ■
  └─ normalizeNumberedLists()    → Standardize all numbered lists to 1. a. i.

PHASE 2: STANDARDIZE FORMATTING
  ├─ standardizeBulletSymbols()        → Format bullets: Arial 12pt bold black
  └─ standardizeNumberedListPrefixes() → Format prefixes: Verdana 12pt bold black

PHASE 3: NORMALIZE INDENTATION
  └─ normalizeAllListIndentation() → Consistent 0.5" increments, 0.25" hanging

================================================================================
BEGINNING EXECUTION
================================================================================

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SCENARIO 1: Normalize All Lists
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Description: Apply all three phases to standardize all list formatting

Loading: Test_Code.docx
✓ Document loaded successfully

--- PHASE 1: Normalize List Types ---
✓ Normalized 45 bullet list items
  - Unified to standard bullet symbols: • (level 0), ○ (level 1), ■ (level 2)
✓ Normalized 32 numbered list items
  - Unified to standard formats: 1. (decimal), a. (lowerLetter), i. (lowerRoman)

--- PHASE 2: Standardize Formatting ---
✓ Standardized 3 bullet lists
  - 27 levels formatted: Arial 12pt bold black
✓ Standardized 2 numbered lists
  - 18 levels formatted: Verdana 12pt bold black

--- PHASE 3: Normalize Indentation ---
✓ Normalized indentation for 5 list definitions
  - Left indent: 720 twips × (level + 1) = 0.5" increments
  - Hanging indent: 360 twips = 0.25" for all levels

✓ Scenario 1 Complete: All lists normalized and standardized

Saving: examples/output/test-list-formatted-scenario1.docx
✓ Document saved successfully

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✓ SCENARIO 1 COMPLETED SUCCESSFULLY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

## API Reference

### Phase 1 Methods

- **[`doc.normalizeBulletLists()`](../src/core/Document.ts:2487)** - Returns count of bullet items normalized
- **[`doc.normalizeNumberedLists()`](../src/core/Document.ts:2421)** - Returns count of numbered items normalized

### Phase 2 Methods

- **[`doc.standardizeBulletSymbols(options?)`](../src/core/Document.ts:4528)** - Returns `{ listsUpdated, levelsModified }`

  - `options.font` - Font name (default: 'Arial')
  - `options.fontSize` - Font size in half-points (default: 24 = 12pt)
  - `options.bold` - Bold formatting (default: true)
  - `options.color` - Hex color without # (default: '000000')

- **[`doc.standardizeNumberedListPrefixes(options?)`](../src/core/Document.ts:4607)** - Returns `{ listsUpdated, levelsModified }`
  - `options.font` - Font name (default: 'Verdana')
  - `options.fontSize` - Font size in half-points (default: 24 = 12pt)
  - `options.bold` - Bold formatting (default: true)
  - `options.color` - Hex color without # (default: '000000')

### Phase 3 Methods

- **[`doc.normalizeAllListIndentation()`](../src/core/Document.ts:2834)** - Returns count of list instances normalized
  - Applies standard indentation: `leftIndent = 720 × (level + 1)` twips
  - Applies standard hanging indent: `360` twips (0.25")

## Requirements

- Node.js 14+
- TypeScript 4+
- `Test_Code.docx` in project root (source document with lists)

## Related Documentation

- [Document API Reference](../src/core/Document.ts)
- [NumberingManager API](../src/formatting/NumberingManager.ts)
- [Bullet Customization Example](./bullet-customization-example.ts)
