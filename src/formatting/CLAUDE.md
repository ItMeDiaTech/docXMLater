# Formatting Module Documentation

The `formatting` module contains classes for managing styles, numbering systems, and document formatting.

## Module Overview

The formatting module provides comprehensive support for Word document formatting, including custom styles, multi-level lists, and style inheritance.

**Location:** `src/formatting/`

**Key Classes:**
- `Style` - Represents a single style definition
- `StylesManager` - Manages all styles in a document
- `NumberingLevel` - Represents a single numbering level
- `AbstractNumbering` - Represents abstract numbering definition
- `NumberingInstance` - Represents concrete numbering instance
- `NumberingManager` - Manages all numbering in a document

## Architecture

### Style System

#### Style Class

**File:** `src/formatting/Style.ts`

Represents a style definition in a Word document.

**Style Types:**
- `paragraph` - Paragraph style (e.g., Heading1, Normal)
- `character` - Character/run style (e.g., Emphasis, Strong)
- `table` - Table style (e.g., TableGrid, LightShading)
- `numbering` - Numbering style (linked to list definitions)

**Properties:**
```typescript
interface StyleProperties {
  styleId: string;          // Unique ID (e.g., "Heading1")
  type: StyleType;          // Style type
  name?: string;            // Display name
  basedOn?: string;         // Parent style ID
  next?: string;            // Next paragraph style
  link?: string;            // Linked style ID
  runFormatting?: RunFormatting;       // Character formatting
  paragraphFormatting?: ParagraphFormatting;  // Paragraph formatting
  tableFormatting?: TableFormatting;   // Table formatting
}
```

**Style Inheritance:**
Styles can inherit from other styles using `basedOn`:
```typescript
const heading2 = new Style('Heading2', 'paragraph');
heading2.setBasedOn('Heading1');  // Inherits from Heading1
heading2.setRunFormatting({
  fontSize: 16  // Override font size
});
// Result: Heading2 has all Heading1 formatting except fontSize
```

**Built-in Styles:**
- `Normal` - Default paragraph style
- `Heading1` through `Heading6` - Heading styles
- `Title` - Document title
- `Subtitle` - Document subtitle
- `DefaultParagraphFont` - Default character style
- `TableGrid` - Default table style

#### StylesManager Class

**File:** `src/formatting/StylesManager.ts`

Manages all styles in a document.

**Core Responsibilities:**
1. Store and retrieve styles
2. Validate style definitions
3. Generate styles.xml
4. Parse styles from existing documents
5. Ensure built-in styles exist

**Style Resolution:**
When applying a style, the manager resolves the full formatting by:
1. Start with style's own formatting
2. Merge with basedOn style formatting (recursive)
3. Apply default formatting for style type
4. Return final computed formatting

**Validation:**
The StylesManager validates:
- Style IDs are unique
- basedOn references exist
- No circular inheritance
- Required properties are set
- Format values are valid

**XML Generation:**
Generates `word/styles.xml` with proper structure:
```xml
<w:styles>
  <w:docDefaults>...</w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr>...</w:pPr>
    <w:rPr>...</w:rPr>
  </w:style>
  <!-- More styles... -->
</w:styles>
```

### Numbering System

The numbering system handles multi-level lists (numbered and bulleted).

#### Concepts

**Abstract Numbering:**
- Defines the structure of a list (9 levels max)
- Specifies format for each level (decimal, roman, bullet, etc.)
- Defines indentation and alignment
- Reusable template

**Numbering Instance:**
- Links an abstract numbering to actual use
- Can override levels from abstract definition
- Referenced by paragraphs via numId

**Numbering Level:**
- Defines format for one level of a list
- Properties: format, text, alignment, indentation, font

#### NumberingLevel Class

**File:** `src/formatting/NumberingLevel.ts`

Represents a single level in a numbering definition.

**Number Formats:**
- `decimal` - 1, 2, 3, 4...
- `upperRoman` - I, II, III, IV...
- `lowerRoman` - i, ii, iii, iv...
- `upperLetter` - A, B, C, D...
- `lowerLetter` - a, b, c, d...
- `bullet` - •, ○, ▪, etc.

**Level Text:**
Defines how numbers are displayed:
- `%1.` - Level 1 number with period (e.g., "1.")
- `%1.%2.` - Level 1 and 2 (e.g., "1.1.")
- `%1.%2.%3.` - Three levels (e.g., "1.1.1.")
- `•` - Bullet character

**Indentation:**
Standard Word indentation formula:
```
leftIndent = 720 + (level * 360)  // in twips
hangingIndent = 360                // in twips

Level 0: 720 twips (0.5 inch)
Level 1: 1080 twips (0.75 inch)
Level 2: 1440 twips (1.0 inch)
Level 3: 1800 twips (1.25 inch)
```

**Helper Methods:**
```typescript
// Create standard bullet level
const level = NumberingLevel.createBulletLevel(0, '•');

// Create numbered level
const level = NumberingLevel.createNumberedLevel(0, 'decimal', '%1.');

// Calculate standard indentation
const indent = NumberingLevel.calculateStandardIndentation(level);

// Get standard number format for level
const format = NumberingLevel.getStandardNumberFormat(level);

// Get bullet symbol with font
const { symbol, font } = NumberingLevel.getBulletSymbolWithFont(level, 'standard');
```

#### AbstractNumbering Class

**File:** `src/formatting/AbstractNumbering.ts`

Represents an abstract numbering definition (template for lists).

**Properties:**
```typescript
interface AbstractNumberingProperties {
  abstractNumId: number;    // Unique ID (1-based)
  levels: NumberingLevel[]; // Up to 9 levels
  multiLevel?: boolean;     // Multi-level list?
  name?: string;            // Display name
}
```

**Factory Methods:**
```typescript
// Create numbered list (1., 2., 3.)
const abstractNum = AbstractNumbering.createNumberedList(1, 9);

// Create bulleted list (•, ○, ▪)
const abstractNum = AbstractNumbering.createBulletedList(1, 9);

// Create custom list
const abstractNum = new AbstractNumbering(1);
abstractNum.addLevel(NumberingLevel.createNumberedLevel(0, 'decimal', '%1.'));
abstractNum.addLevel(NumberingLevel.createNumberedLevel(1, 'lowerLetter', '%2.'));
```

**Multi-Level List Pattern:**
```
Level 0: 1., 2., 3.           (decimal)
Level 1: a., b., c.           (lowerLetter)
Level 2: i., ii., iii.        (lowerRoman)
Level 3: A., B., C.           (upperLetter)
Level 4: I., II., III.        (upperRoman)
Level 5+: cycles back to decimal
```

#### NumberingInstance Class

**File:** `src/formatting/NumberingInstance.ts`

Represents a concrete numbering instance.

**Properties:**
```typescript
interface NumberingInstanceProperties {
  numId: number;                 // Unique ID (1-based)
  abstractNumId: number;         // Links to AbstractNumbering
  levelOverrides?: Map<number, NumberingLevel>;  // Level overrides
}
```

**Level Overrides:**
Instances can override specific levels from the abstract numbering:
```typescript
const instance = new NumberingInstance(1, 1);  // numId=1, abstractNumId=1
const override = NumberingLevel.createNumberedLevel(0, 'upperRoman', '%1.');
instance.addLevelOverride(0, override);
// Result: Level 0 uses Roman numerals instead of decimal
```

#### NumberingManager Class

**File:** `src/formatting/NumberingManager.ts`

Manages all numbering in a document.

**Core Responsibilities:**
1. Store abstract numbering definitions
2. Store numbering instances
3. Generate numbering.xml
4. Parse numbering from existing documents
5. Assign unique IDs

**Usage Pattern:**
```typescript
// Create abstract numbering
const abstractNum = AbstractNumbering.createNumberedList(1, 9);
manager.addAbstractNumbering(abstractNum);

// Create instance
const instance = new NumberingInstance(1, 1);
manager.addNumberingInstance(instance);

// Apply to paragraph
paragraph.setNumbering(1, 0);  // numId=1, level=0
```

**XML Generation:**
Generates `word/numbering.xml`:
```xml
<w:numbering>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <!-- More levels... -->
  </w:abstractNum>

  <w:num w:numId="1">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>
```

## Data Structures

### Style Formatting Objects

**RunFormatting (Character Formatting):**
```typescript
interface RunFormatting {
  bold?: boolean;
  italic?: boolean;
  underline?: string;         // 'single', 'double', etc.
  strikethrough?: boolean;
  fontSize?: number;          // in points
  font?: string;
  color?: string;             // hex color
  highlight?: string;         // highlight color
  subscript?: boolean;
  superscript?: boolean;
  smallCaps?: boolean;
  allCaps?: boolean;
}
```

**ParagraphFormatting:**
```typescript
interface ParagraphFormatting {
  alignment?: ParagraphAlignment;  // 'left', 'center', 'right', 'justify'
  leftIndent?: number;             // in twips
  rightIndent?: number;
  firstLineIndent?: number;
  hangingIndent?: number;
  spacingBefore?: number;          // in twips
  spacingAfter?: number;
  lineSpacing?: number;
  borders?: ParagraphBorders;
  shading?: string;                // hex color
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
}
```

### Inheritance Chain Example

```
Style: Heading2
  └─ basedOn: Heading1
      └─ basedOn: Normal
          └─ (default formatting)

Resolved formatting for Heading2:
1. Start with default paragraph formatting
2. Merge Normal formatting
3. Merge Heading1 formatting
4. Merge Heading2 formatting
5. Result: Final computed style
```

## Best Practices

### Creating Custom Styles

```typescript
// Create custom paragraph style
const customStyle = new Style('CustomHeading', 'paragraph');
customStyle.setName('Custom Heading');
customStyle.setBasedOn('Heading1');  // Inherit from built-in
customStyle.setRunFormatting({
  color: '0070C0',
  fontSize: 18
});
customStyle.setParagraphFormatting({
  spacingBefore: 240,
  spacingAfter: 120
});

doc.getStylesManager().addStyle(customStyle);
```

### Creating Multi-Level Lists

```typescript
// Create custom numbered list with specific formats
const abstractNum = new AbstractNumbering(1);

// Level 0: 1., 2., 3.
abstractNum.addLevel(
  NumberingLevel.createNumberedLevel(0, 'decimal', '%1.')
);

// Level 1: a., b., c.
abstractNum.addLevel(
  NumberingLevel.createNumberedLevel(1, 'lowerLetter', '%2.')
);

// Level 2: i., ii., iii.
abstractNum.addLevel(
  NumberingLevel.createNumberedLevel(2, 'lowerRoman', '%3.')
);

doc.getNumberingManager().addAbstractNumbering(abstractNum);

const instance = new NumberingInstance(1, 1);
doc.getNumberingManager().addNumberingInstance(instance);
```

### Applying Numbering to Paragraphs

```typescript
const para1 = doc.createParagraph();
para1.addText('First item');
para1.setNumbering(1, 0);  // numId=1, level=0

const para2 = doc.createParagraph();
para2.addText('Second item');
para2.setNumbering(1, 0);  // Same list continues

const para3 = doc.createParagraph();
para3.addText('Sub-item');
para3.setNumbering(1, 1);  // Level 1 (indented)
```

### Using Built-in Bullet Styles

```typescript
// Standard bullet: •
const level1 = NumberingLevel.createBulletLevel(0, '•');

// Circle bullet: ○
const level2 = NumberingLevel.createBulletLevel(1, '○');

// Square bullet: ▪
const level3 = NumberingLevel.createBulletLevel(2, '▪');

// Helper method for standard styles
const { symbol, font } = NumberingLevel.getBulletSymbolWithFont(0, 'standard');
// Returns: { symbol: '•', font: 'Calibri' }
```

## Testing

The formatting module has comprehensive test coverage:

**File:** `tests/formatting/StyleValidation.test.ts`
- Style validation (25+ tests)
- Circular reference detection (10+ tests)

**File:** `tests/formatting/StyleEnhancements.test.ts`
- Style inheritance (20+ tests)
- Format merging (15+ tests)

**File:** `tests/formatting/StylesRoundTrip.test.ts`
- Parse and generate consistency (30+ tests)

**File:** `tests/formatting/Numbering.test.ts`
- List creation (40+ tests)
- Level formatting (25+ tests)
- Indentation calculations (15+ tests)

**File:** `tests/formatting/TableStyles.test.ts`
- Table style application (20+ tests)

**Total: 200+ tests covering formatting functionality**

## Common Patterns

### Pattern 1: Corporate Document Template

```typescript
// Define corporate styles
const titleStyle = new Style('CorporateTitle', 'paragraph');
titleStyle.setName('Corporate Title');
titleStyle.setRunFormatting({
  font: 'Arial',
  fontSize: 24,
  bold: true,
  color: '003366'
});
titleStyle.setParagraphFormatting({
  alignment: 'center',
  spacingAfter: 240
});

const bodyStyle = new Style('CorporateBody', 'paragraph');
bodyStyle.setName('Corporate Body');
bodyStyle.setRunFormatting({
  font: 'Arial',
  fontSize: 11
});
bodyStyle.setParagraphFormatting({
  alignment: 'justify',
  lineSpacing: 276  // 1.15 line spacing
});

doc.getStylesManager().addStyle(titleStyle);
doc.getStylesManager().addStyle(bodyStyle);
```

### Pattern 2: Academic Paper Numbering

```typescript
// Create academic numbering (I, A, 1, a, i)
const abstractNum = new AbstractNumbering(1);
abstractNum.addLevel(NumberingLevel.createNumberedLevel(0, 'upperRoman', '%1.'));
abstractNum.addLevel(NumberingLevel.createNumberedLevel(1, 'upperLetter', '%2.'));
abstractNum.addLevel(NumberingLevel.createNumberedLevel(2, 'decimal', '%3.'));
abstractNum.addLevel(NumberingLevel.createNumberedLevel(3, 'lowerLetter', '%4.'));
abstractNum.addLevel(NumberingLevel.createNumberedLevel(4, 'lowerRoman', '%5.'));

doc.getNumberingManager().addAbstractNumbering(abstractNum);
```

### Pattern 3: Custom Bullet List

```typescript
// Create custom bullet list with specific symbols
const abstractNum = new AbstractNumbering(1);

// Level 0: ➢ (arrow)
const level0 = NumberingLevel.createBulletLevel(0, '➢');
level0.setFont('Wingdings');
abstractNum.addLevel(level0);

// Level 1: ✓ (check mark)
const level1 = NumberingLevel.createBulletLevel(1, '✓');
level1.setFont('Wingdings');
abstractNum.addLevel(level1);

doc.getNumberingManager().addAbstractNumbering(abstractNum);
```

## Performance Considerations

- **Style Resolution**: Cached to avoid repeated computation
- **Numbering Generation**: Levels generated on-demand
- **XML Generation**: Minimal templates, no redundant elements
- **Validation**: Performed once on add, not on every access

## Troubleshooting

### Issue: Circular Style Inheritance

**Problem:** Style A based on B, B based on A
**Solution:** StylesManager detects and prevents circular references
**Error:** `Circular style reference detected: A -> B -> A`

### Issue: Missing Abstract Numbering

**Problem:** Paragraph references numId that doesn't exist
**Solution:** Ensure abstract numbering and instance are added to manager
**Fix:**
```typescript
manager.addAbstractNumbering(abstractNum);
manager.addNumberingInstance(instance);
paragraph.setNumbering(instance.getNumId(), 0);
```

### Issue: Incorrect Indentation

**Problem:** List levels not indenting correctly
**Solution:** Use standard indentation calculation
**Fix:**
```typescript
const indent = NumberingLevel.calculateStandardIndentation(level);
level.setLeftIndent(indent.left);
level.setHangingIndent(indent.hanging);
```

## See Also

- `src/core/CLAUDE.md` - Core document classes
- `src/elements/CLAUDE.md` - Document elements
- `src/xml/CLAUDE.md` - XML generation
