# API Refactor Plan for v4.0.0

## Breaking Changes: Style Management Overhaul

### 🎯 Goals

1. **Remove** broken `applyCustomStylesToDocument()` method
2. **Rename** `applyCustomFormattingToExistingStyles()` → `applyStyles()`
3. **Enhance** StylesManager with helper methods
4. **Mandate** StylesManager usage for style operations
5. **Simplify** API surface for common use cases

---

## Current State Analysis

### Existing StylesManager Capabilities ✅

Already available in [`StylesManager.ts`](../src/formatting/StylesManager.ts:1):

**Basic Operations:**

- [`addStyle()`](../src/formatting/StylesManager.ts:97) - Add/update style
- [`getStyle()`](../src/formatting/StylesManager.ts:108) - Retrieve style by ID
- [`hasStyle()`](../src/formatting/StylesManager.ts:119) - Check existence
- [`removeStyle()`](../src/formatting/StylesManager.ts:137) - Delete style
- [`getAllStyles()`](../src/formatting/StylesManager.ts:145) - Get all styles
- [`clear()`](../src/formatting/StylesManager.ts:259) - Clear all styles

**Style Creation:**

- [`createParagraphStyle()`](../src/formatting/StylesManager.ts:309)
- [`createCharacterStyle()`](../src/formatting/StylesManager.ts:328)
- [`createTableStyle()`](../src/formatting/StylesManager.ts:235)

**Filtering:**

- [`getStylesByType()`](../src/formatting/StylesManager.ts:154) - Filter by type
- [`getQuickStyles()`](../src/formatting/StylesManager.ts:163) - Gallery styles
- [`getVisibleStyles()`](../src/formatting/StylesManager.ts:176) - Non-hidden
- [`getStylesByPriority()`](../src/formatting/StylesManager.ts:189) - UI order
- [`getTableStyles()`](../src/formatting/StylesManager.ts:224) - Table styles only

**Analysis:**

- [`getLinkedStyle()`](../src/formatting/StylesManager.ts:206) - Get linked style
- [`getStats()`](../src/formatting/StylesManager.ts:285) - Statistics
- [`validate()`](../src/formatting/StylesManager.ts:434) - XML validation

### Access Pattern

Already exposed via Document:

```typescript
const stylesManager = doc.getStylesManager();
stylesManager.addStyle(myStyle);
```

---

## Proposed Changes for v4.0.0

### 1. Method Removal & Rename

#### In [`Document.ts`](../src/core/Document.ts:1):

**Remove (Lines 3103-3181):**

```typescript
❌ DELETE: applyCustomStylesToDocument()
```

**Rename (Line 3311):**

```typescript
// OLD
applyCustomFormattingToExistingStyles(options?: ApplyCustomFormattingOptions)

// NEW
applyStyles(options?: ApplyStylesOptions)  // Shorter, clearer name
```

**Add Alias for Migration (temporary):**

```typescript
/**
 * @deprecated Use {@link applyStyles} instead (renamed in v4.0.0)
 */
applyCustomFormattingToExistingStyles(options?: ApplyStylesOptions) {
  return this.applyStyles(options);
}
```

### 2. Enhanced StylesManager Methods

Add to [`StylesManager.ts`](../src/formatting/StylesManager.ts:556):

#### A. Bulk Style Operations

```typescript
/**
 * Applies formatting to multiple styles at once
 * @param configs - Map of styleId to formatting config
 * @returns Count of styles updated
 */
applyBulkFormatting(configs: Map<string, StyleConfig>): number {
  let count = 0;
  for (const [styleId, config] of configs) {
    const style = this.getStyle(styleId);
    if (style) {
      if (config.run) style.setRunFormatting(config.run);
      if (config.paragraph) style.setParagraphFormatting(config.paragraph);
      count++;
    }
  }
  return count;
}
```

#### B. Style Search & Filter

```typescript
/**
 * Searches styles by name (case-insensitive)
 * @param searchTerm - Text to search for
 * @returns Matching styles
 */
searchByName(searchTerm: string): Style[] {
  const term = searchTerm.toLowerCase();
  return this.getAllStyles().filter(s =>
    s.getName().toLowerCase().includes(term)
  );
}

/**
 * Finds styles using a specific font
 * @param fontName - Font family name
 * @returns Styles using this font
 */
findByFont(fontName: string): Style[] {
  return this.getAllStyles().filter(s => {
    const run = s.getRunFormatting();
    return run?.font === fontName;
  });
}

/**
 * Finds styles with specific properties
 * @param predicate - Filter function
 * @returns Matching styles
 */
findStyles(predicate: (style: Style) => boolean): Style[] {
  return this.getAllStyles().filter(predicate);
}
```

#### C. Style Comparison

```typescript
/**
 * Compares two styles and returns differences
 * @param styleId1 - First style ID
 * @param styleId2 - Second style ID
 * @returns Object describing differences
 */
compareStyles(styleId1: string, styleId2: string): {
  runDifferences: Record<string, { style1: any; style2: any }>;
  paragraphDifferences: Record<string, { style1: any; style2: any }>;
  identical: boolean;
} {
  const style1 = this.getStyle(styleId1);
  const style2 = this.getStyle(styleId2);

  if (!style1 || !style2) {
    throw new Error('One or both styles not found');
  }

  // Implementation compares run and paragraph formatting
  // Returns differences object
}
```

#### D. Style Templates

```typescript
/**
 * Creates a style template set for common document types
 * @param template - Template type
 * @returns Array of created styles
 */
applyStyleTemplate(template: 'academic' | 'business' | 'report' | 'apa' | 'mla'): Style[] {
  const styles: Style[] = [];

  switch (template) {
    case 'academic':
      // Times New Roman 12pt, double spacing
      styles.push(this.createAcademicStyles());
      break;
    case 'business':
      // Arial 11pt, single spacing
      styles.push(this.createBusinessStyles());
      break;
    // ... etc
  }

  for (const style of styles) {
    this.addStyle(style);
  }

  return styles;
}

private createAcademicStyles(): Style[] {
  return [
    Style.create({
      styleId: 'Normal',
      name: 'Normal',
      type: 'paragraph',
      runFormatting: { font: 'Times New Roman', size: 12 },
      paragraphFormatting: {
        spacing: { line: 480, lineRule: 'auto' }, // Double spacing
        alignment: 'justify'
      }
    }),
    // ... Heading1-3 with Times New Roman
  ];
}
```

#### E. Style Inheritance Analysis

```typescript
/**
 * Gets the complete inheritance chain for a style
 * @param styleId - Style ID to analyze
 * @returns Array of styles from base to derived
 */
getInheritanceChain(styleId: string): Style[] {
  const chain: Style[] = [];
  let current = this.getStyle(styleId);

  while (current) {
    chain.unshift(current);
    const props = current.getProperties();
    current = props.basedOn ? this.getStyle(props.basedOn) : undefined;
  }

  return chain;
}

/**
 * Gets all styles derived from a base style
 * @param baseStyleId - Base style ID
 * @returns Styles that inherit from this one
 */
getDerivedStyles(baseStyleId: string): Style[] {
  return this.getAllStyles().filter(s =>
    s.getProperties().basedOn === baseStyleId
  );
}
```

#### F. Style Export/Import

```typescript
/**
 * Exports style as JSON for reuse
 * @param styleId - Style to export
 * @returns JSON representation
 */
exportStyle(styleId: string): string {
  const style = this.getStyle(styleId);
  if (!style) {
    throw new Error(`Style ${styleId} not found`);
  }
  return JSON.stringify(style.getProperties(), null, 2);
}

/**
 * Imports style from JSON
 * @param json - JSON style definition
 * @returns Created style
 */
importStyle(json: string): Style {
  const props = JSON.parse(json);
  const style = Style.create(props);
  this.addStyle(style);
  return style;
}

/**
 * Exports all styles as JSON
 * @returns JSON array of all styles
 */
exportAllStyles(): string {
  const styles = this.getAllStyles().map(s => s.getProperties());
  return JSON.stringify(styles, null, 2);
}

/**
 * Imports multiple styles from JSON array
 * @param json - JSON array of style definitions
 * @returns Array of created styles
 */
importStyles(json: string): Style[] {
  const propsArray = JSON.parse(json);
  const styles: Style[] = [];

  for (const props of propsArray) {
    const style = Style.create(props);
    this.addStyle(style);
    styles.push(style);
  }

  return styles;
}
```

#### G. Style Validation & Cleanup

```typescript
/**
 * Finds unused styles (not referenced by any paragraphs)
 * @param paragraphs - All paragraphs in document
 * @returns Array of unused style IDs
 */
findUnusedStyles(paragraphs: Paragraph[]): string[] {
  const usedStyles = new Set<string>();

  for (const para of paragraphs) {
    const style = para.getStyle();
    if (style) usedStyles.add(style);
  }

  const allStyleIds = this.getAllStyles().map(s => s.getStyleId());
  return allStyleIds.filter(id => !usedStyles.has(id));
}

/**
 * Removes all unused styles
 * @param paragraphs - All paragraphs in document
 * @returns Count of removed styles
 */
cleanupUnusedStyles(paragraphs: Paragraph[]): number {
  const unused = this.findUnusedStyles(paragraphs);
  let count = 0;

  for (const styleId of unused) {
    // Don't remove built-in styles
    if (!StylesManager.isBuiltInStyle(styleId)) {
      if (this.removeStyle(styleId)) count++;
    }
  }

  return count;
}

/**
 * Validates all style references are resolvable
 * @returns Validation result with broken references
 */
validateStyleReferences(): {
  valid: boolean;
  brokenReferences: Array<{ styleId: string; basedOn: string }>;
  circularReferences: Array<string[]>;
} {
  const broken: Array<{ styleId: string; basedOn: string }> = [];
  const circular: Array<string[]> = [];

  for (const style of this.getAllStyles()) {
    const props = style.getProperties();

    // Check basedOn references exist
    if (props.basedOn && !this.hasStyle(props.basedOn)) {
      broken.push({ styleId: props.styleId, basedOn: props.basedOn });
    }

    // Check for circular references
    const chain = this.getInheritanceChain(props.styleId);
    const ids = chain.map(s => s.getStyleId());
    if (new Set(ids).size !== ids.length) {
      circular.push(ids);
    }
  }

  return {
    valid: broken.length === 0 && circular.length === 0,
    brokenReferences: broken,
    circularReferences: circular
  };
}
```

#### H. Quick Style Manipulation

```typescript
/**
 * Sets a style as Quick Style (shows in gallery)
 * @param styleId - Style to make quick
 * @returns This manager for chaining
 */
setAsQuickStyle(styleId: string): this {
  const style = this.getStyle(styleId);
  if (style) {
    const props = style.getProperties();
    props.qFormat = true;
    props.semiHidden = false;
    style.setProperties(props);
  }
  return this;
}

/**
 * Hides a style from the gallery
 * @param styleId - Style to hide
 * @returns This manager for chaining
 */
hideFromGallery(styleId: string): this {
  const style = this.getStyle(styleId);
  if (style) {
    const props = style.getProperties();
    props.semiHidden = true;
    style.setProperties(props);
  }
  return this;
}

/**
 * Sets UI priority for style ordering
 * @param styleId - Style ID
 * @param priority - Priority (lower = higher in list)
 * @returns This manager for chaining
 */
setStylePriority(styleId: string, priority: number): this {
  const style = this.getStyle(styleId);
  if (style) {
    const props = style.getProperties();
    props.uiPriority = priority;
    style.setProperties(props);
  }
  return this;
}
```

---

## New API Design

### Simplified Document API

```typescript
// In Document.ts

/**
 * Applies styles to document with custom formatting
 * Replaces: applyCustomFormattingToExistingStyles()
 *
 * @param options - Style formatting options
 * @returns Result object with success flags per style
 */
public applyStyles(options?: ApplyStylesOptions): ApplyStylesResult {
  // Same implementation as applyCustomFormattingToExistingStyles
  // Just with cleaner name
}

/**
 * Gets the styles manager for advanced style operations
 * Encourages using StylesManager directly for complex scenarios
 *
 * @returns StylesManager instance
 */
public styles(): StylesManager {
  return this.getStylesManager();
}
```

### Usage Examples - New API

#### Simple Use Case

```typescript
// Quick style application
doc.applyStyles({
  normal: {
    run: { font: "Arial", size: 11, preserveBold: true },
  },
});
```

#### Advanced Use Case (via StylesManager)

```typescript
// Access StylesManager
const styles = doc.styles();

// Template application
styles.applyStyleTemplate("business");

// Bulk font change
const arialStyles = styles.findByFont("Calibri");
for (const style of arialStyles) {
  style.setRunFormatting({ ...style.getRunFormatting(), font: "Arial" });
}

// Cleanup
const removed = styles.cleanupUnusedStyles(doc.getAllParagraphs());

// Export for reuse
const styleJson = styles.exportAllStyles();
await fs.writeFile("my-styles.json", styleJson);
```

---

## StylesManager Enhancements Summary

### New Methods to Add

| Category        | Method                      | Purpose                           |
| --------------- | --------------------------- | --------------------------------- |
| **Bulk Ops**    | `applyBulkFormatting()`     | Update multiple styles at once    |
| **Search**      | `searchByName()`            | Find styles by name               |
| **Search**      | `findByFont()`              | Find styles using specific font   |
| **Search**      | `findStyles()`              | Generic predicate-based search    |
| **Compare**     | `compareStyles()`           | Show differences between styles   |
| **Templates**   | `applyStyleTemplate()`      | Apply preset style sets           |
| **Templates**   | `createAcademicStyles()`    | Academic paper template (private) |
| **Templates**   | `createBusinessStyles()`    | Business doc template (private)   |
| **Templates**   | `createReportStyles()`      | Report template (private)         |
| **Inheritance** | `getInheritanceChain()`     | Full inheritance path             |
| **Inheritance** | `getDerivedStyles()`        | Children of a style               |
| **Export**      | `exportStyle()`             | JSON export single style          |
| **Export**      | `importStyle()`             | JSON import single style          |
| **Export**      | `exportAllStyles()`         | JSON export all styles            |
| **Export**      | `importStyles()`            | JSON import multiple styles       |
| **Validation**  | `findUnusedStyles()`        | Detect orphaned styles            |
| **Validation**  | `cleanupUnusedStyles()`     | Remove orphaned styles            |
| **Validation**  | `validateStyleReferences()` | Check broken references           |
| **Gallery**     | `setAsQuickStyle()`         | Show in Word gallery              |
| **Gallery**     | `hideFromGallery()`         | Hide from gallery                 |
| **Gallery**     | `setStylePriority()`        | Set UI ordering                   |

---

## Type Definitions

### Rename Interface

```typescript
// OLD (styleConfig.ts)
export interface ApplyCustomFormattingOptions { ... }

// NEW (styleConfig.ts - v4.0.0)
export interface ApplyStylesOptions {
  heading1?: StyleConfig;
  heading2?: Heading2Config;
  heading3?: StyleConfig;
  normal?: StyleConfig;
  listParagraph?: StyleConfig;
  preserveBlankLinesAfterHeader2Tables?: boolean;
}

// Keep old name as alias for migration
/** @deprecated Use ApplyStylesOptions instead */
export type ApplyCustomFormattingOptions = ApplyStylesOptions;
```

### New Result Type

```typescript
export interface ApplyStylesResult {
  heading1: boolean;
  heading2: boolean;
  heading3: boolean;
  normal: boolean;
  listParagraph: boolean;
}
```

---

## Migration Strategy for v4.0.0

### Phase 1: Deprecation Warnings (v3.6.0)

- Add runtime warnings to old methods
- Update documentation

### Phase 2: New Methods (v3.7.0)

- Add `applyStyles()` as alias
- Add enhanced StylesManager methods
- Keep both old and new methods

### Phase 3: Breaking Changes (v4.0.0)

- **REMOVE:** `applyCustomStylesToDocument()`
- **RENAME:** `applyCustomFormattingToExistingStyles()` → `applyStyles()`
- Keep compatibility alias with deprecation warning

### Phase 4: Full Migration (v4.1.0)

- Remove compatibility alias
- Pure new API

---

## Documentation Updates

### 1. Update README.md

````markdown
## Styling Documents

### Quick Styling with applyStyles()

```typescript
const doc = Document.create();

// Apply standard styles
doc.applyStyles({
  normal: {
    run: { font: "Arial", size: 11, preserveBold: true },
  },
  heading1: {
    run: { font: "Arial", size: 18, bold: true },
  },
});
```
````

### Advanced Styling with StylesManager

```typescript
// Access style manager
const styles = doc.styles();

// Apply template
styles.applyStyleTemplate("business");

// Find and modify
const headings = styles.findStyles((s) => s.getStyleId().startsWith("Heading"));

// Export styles
const json = styles.exportAllStyles();
```

````

### 2. Create Examples

Create [`examples/styles/advanced-style-management.ts`](../examples/styles/advanced-style-management.ts:1):
- StylesManager template usage
- Bulk operations
- Style search/filter
- Export/import
- Validation

---

## Breaking Change Communication

### CHANGELOG.md Entry

```markdown
## [4.0.0] - YYYY-MM-DD

### 💥 BREAKING CHANGES

#### Removed Methods
- **`applyCustomStylesToDocument()`** - Removed (was broken, created undefined styles)
  - **Migration:** Use `applyStyles()` instead

#### Renamed Methods
- **`applyCustomFormattingToExistingStyles()`** → **`applyStyles()`**
  - Shorter, clearer name
  - Same functionality, better API
  - Temporary alias provided with deprecation warning

### ✨ Added

#### StylesManager Enhancements
- `applyBulkFormatting()` - Bulk style updates
- `searchByName()` - Search styles
- `findByFont()` - Find styles by font
- `compareStyles()` - Compare two styles
- `applyStyleTemplate()` - Apply preset templates (academic, business, report, APA, MLA)
- `getInheritanceChain()` - Analyze style inheritance
- `getDerivedStyles()` - Find child styles
- `exportStyle()` / `importStyle()` - JSON serialization
- `findUnusedStyles()` / `cleanupUnusedStyles()` - Orphan detection
- `validateStyleReferences()` - Reference validation
- `setAsQuickStyle()` / `hideFromGallery()` - Gallery management

#### New Convenience Methods
- `doc.styles()` - Shorter accessor for StylesManager

### 🔄 Migration Guide

See [`docs/guides/v4-migration-guide.md`](./guides/v4-migration-guide.md)
````

---

## Implementation Priority

### High Priority (v4.0.0)

1. ✅ Remove `applyCustomStylesToDocument()`
2. ✅ Rename to `applyStyles()`
3. ✅ Add `doc.styles()` shortcut
4. ✅ Add type alias for backwards compat

### Medium Priority (v4.0.0 or v4.1.0)

5. Add search methods (`searchByName`, `findByFont`, `findStyles`)
6. Add validation methods (`findUnusedStyles`, `validateStyleReferences`)
7. Add inheritance methods (`getInheritanceChain`, `getDerivedStyles`)

### Low Priority (v4.1.0+)

8. Add template system (`applyStyleTemplate`)
9. Add export/import (`exportStyle`, `importStyle`)
10. Add comparison (`compareStyles`)
11. Add gallery helpers (`setAsQuickStyle`, `hideFromGallery`)

---

## File Structure

```
src/
├── formatting/
│   ├── StylesManager.ts         # Enhanced with new methods
│   ├── style-templates/         # NEW - Template definitions
│   │   ├── academic.ts
│   │   ├── business.ts
│   │   ├── report.ts
│   │   ├── apa.ts
│   │   └── mla.ts
│   └── Style.ts                 # Existing
├── types/
│   └── styleConfig.ts           # Rename interface
└── core/
    └── Document.ts              # Remove old, rename new method

docs/
├── guides/
│   ├── style-application-migration.md  # Already created
│   └── v4-migration-guide.md           # NEW - Full v4 migration
└── DEPRECATION_applyCustomStylesToDocument.md  # Archive/remove

examples/
└── styles/
    ├── advanced-style-management.ts    # NEW - StylesManager examples
    └── style-templates.ts              # NEW - Template examples
```

---

## Next Steps

1. **Review & Approve** this refactor plan
2. **Prioritize** which enhancements to include in v4.0.0
3. **Switch to Code Mode** to implement:
   - Remove `applyCustomStylesToDocument()`
   - Rename to `applyStyles()`
   - Add core StylesManager enhancements
   - Create examples
4. **Test** thoroughly
5. **Release** v4.0.0 with breaking changes

---

## Questions for Clarification

1. **Template Priority:** Which document templates are most important? (Academic, Business, APA, MLA, Report?)
2. **StylesManager Scope:** Which helper methods are essential for v4.0.0 vs. can wait for v4.1.0?
3. **Migration Timeline:** How long to keep compatibility alias before final removal?
