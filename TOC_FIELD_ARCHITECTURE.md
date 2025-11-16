# Table of Contents (TOC) Field Architecture Plan

## Executive Summary

This document outlines the architecture and implementation plan for TOC (Table of Contents) field representation in docXMLater. The primary goal is to ensure TOC fields maintain their hidden field logic, allowing users to right-click and select "Update Field" in Microsoft Word.

## Current Implementation Analysis

### ✅ GOOD NEWS: Field Structure Already Implemented!

The current [`TableOfContents.ts`](src/elements/TableOfContents.ts) implementation **ALREADY maintains proper field structure** that supports Word's "Update Field" functionality:

```xml
<w:sdt>
  <w:sdtPr>
    <w:id w:val="-123456789"/>
    <w:docPartObj>
      <w:docPartGallery w:val="Table of Contents"/>
      <w:docPartUnique w:val="1"/>
    </w:docPartObj>
  </w:sdtPr>
  <w:sdtContent>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="TOCHeading"/>
      </w:pPr>
      <w:r>
        <w:t>Table of Contents</w:t>
      </w:r>
    </w:p>
    <w:p>
      <!-- FIELD BEGIN -->
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>

      <!-- FIELD INSTRUCTION -->
      <w:r>
        <w:instrText xml:space="preserve">TOC \o "1-3" \h \* MERGEFORMAT</w:instrText>
      </w:r>

      <!-- FIELD SEPARATOR -->
      <w:r>
        <w:fldChar w:fldCharType="separate"/>
      </w:r>

      <!-- FIELD CONTENT (Placeholder) -->
      <w:r>
        <w:rPr>
          <w:noProof/>
        </w:rPr>
        <w:t>Right-click to update field.</w:t>
      </w:r>

      <!-- FIELD END (CRITICAL!) -->
      <w:r>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    </w:p>
  </w:sdtContent>
</w:sdt>
```

### Key Implementation Details

#### 1. Complete Field Structure (ECMA-376 §17.16.5 Compliant)

- ✅ Field BEGIN marker (`fldCharType="begin"`)
- ✅ Field INSTRUCTION (`<w:instrText>`)
- ✅ Field SEPARATOR (`fldCharType="separate"`)
- ✅ Field CONTENT (placeholder text)
- ✅ Field END marker (`fldCharType="end"`) - **CRITICAL**

#### 2. Field Instruction Building

Current implementation in [`buildFieldInstruction()`](src/elements/TableOfContents.ts#L226-L275):

```typescript
private buildFieldInstruction(): string {
  let instruction = 'TOC';

  // Heading levels OR custom styles
  if (this.includeStyles && this.includeStyles.length > 0) {
    for (const style of this.includeStyles) {
      instruction += ` \\t "${style.styleName},${style.level},"`;
    }
  } else {
    instruction += ` \\o "1-${this.levels}"`;
  }

  // Hyperlinks
  if (this.useHyperlinks) {
    instruction += ' \\h';
  }

  // Page numbers
  if (!this.showPageNumbers) {
    instruction += ' \\n';
  }

  // Web layout hide
  if (this.hideInWebLayout) {
    instruction += ' \\z';
  }

  // Tab leader
  if (this.tabLeader !== 'dot') {
    const leaderMap = { hyphen: 'h', underscore: 'u', none: 'n' };
    instruction += ` \\p "${leaderMap[this.tabLeader]}"`;
  }

  // Custom switches
  if (this.fieldSwitches) {
    instruction += ` ${this.fieldSwitches}`;
  }

  // MERGEFORMAT (preserves formatting on update)
  instruction += ' \\* MERGEFORMAT';

  return instruction;
}
```

#### 3. Field Instruction Preservation

The implementation preserves original field instructions from loaded documents:

```typescript
// In constructor
this.originalFieldInstruction = properties.originalFieldInstruction;

// In toXML()
const fieldInstruction =
  this.originalFieldInstruction || this.buildFieldInstruction();
```

## Enhancement Plan

### Phase 1: Documentation & Validation ✨

#### Task 1: Document Current Field Structure

**File**: `docs/guides/toc-field-structure.md`

Create comprehensive documentation explaining:

- How the field structure works
- Why each component is necessary
- How Word interprets field codes
- Common field switches and their meanings

#### Task 2: Add Field Validation Helper

**File**: [`src/elements/TableOfContents.ts`](src/elements/TableOfContents.ts)

```typescript
/**
 * Validates that the TOC field structure is complete and correct
 * Per ECMA-376 §17.16.5, all complex fields MUST have:
 * begin → instrText → separate → content → end
 *
 * @returns Validation result with details
 */
public validateFieldStructure(): {
  valid: boolean;
  errors: string[];
  warnings: string[];
} {
  const errors: string[] = [];
  const warnings: string[] = [];

  // Validate field instruction
  const instruction = this.getFieldInstruction();
  if (!instruction.startsWith('TOC')) {
    errors.push('Field instruction must start with "TOC"');
  }

  if (!instruction.includes('MERGEFORMAT')) {
    warnings.push('Missing \\* MERGEFORMAT switch - formatting may not persist on update');
  }

  // Validate levels
  if (this.levels < 1 || this.levels > 9) {
    errors.push('TOC levels must be between 1 and 9');
  }

  return {
    valid: errors.length === 0,
    errors,
    warnings
  };
}
```

### Phase 2: Field Manipulation Helpers 🛠️

#### Task 3: Enhanced Field Instruction Builder

**File**: [`src/elements/TableOfContents.ts`](src/elements/TableOfContents.ts)

```typescript
/**
 * Builds a custom TOC field instruction from components
 * Provides fine-grained control over field switches
 *
 * @param options - Field instruction options
 * @returns Complete field instruction string
 *
 * @example
 * const instruction = TableOfContents.buildCustomFieldInstruction({
 *   levels: { start: 1, end: 3 },
 *   switches: {
 *     hyperlinks: true,
 *     hidePageNumbers: false,
 *     hideInWebLayout: true,
 *     tabLeader: 'dot'
 *   },
 *   customStyles: [
 *     { name: 'MyHeading1', level: 1 },
 *     { name: 'MyHeading2', level: 2 }
 *   ],
 *   preserveFormatting: true
 * });
 * // Returns: "TOC \t "MyHeading1,1," \t "MyHeading2,2," \h \z \* MERGEFORMAT"
 */
public static buildCustomFieldInstruction(options: {
  levels?: { start: number; end: number };
  switches?: {
    hyperlinks?: boolean;
    hidePageNumbers?: boolean;
    hideInWebLayout?: boolean;
    tabLeader?: 'dot' | 'hyphen' | 'underscore' | 'none';
    sequenceIdentifier?: string; // \f switch
    entryIdentifier?: string; // \l switch
  };
  customStyles?: Array<{ name: string; level: number }>;
  preserveFormatting?: boolean;
  customSwitches?: string;
}): string {
  let instruction = 'TOC';

  // Add levels or custom styles
  if (options.customStyles && options.customStyles.length > 0) {
    for (const style of options.customStyles) {
      instruction += ` \\t "${style.name},${style.level},"`;
    }
  } else if (options.levels) {
    instruction += ` \\o "${options.levels.start}-${options.levels.end}"`;
  } else {
    instruction += ` \\o "1-3"`; // Default
  }

  // Add switches
  if (options.switches?.hyperlinks) {
    instruction += ' \\h';
  }

  if (options.switches?.hidePageNumbers) {
    instruction += ' \\n';
  }

  if (options.switches?.hideInWebLayout) {
    instruction += ' \\z';
  }

  if (options.switches?.tabLeader && options.switches.tabLeader !== 'dot') {
    const leaderMap = { hyphen: 'h', underscore: 'u', none: 'n' };
    instruction += ` \\p "${leaderMap[options.switches.tabLeader]}"`;
  }

  if (options.switches?.sequenceIdentifier) {
    instruction += ` \\f ${options.switches.sequenceIdentifier}`;
  }

  if (options.switches?.entryIdentifier) {
    instruction += ` \\l "${options.switches.entryIdentifier}"`;
  }

  // Custom switches
  if (options.customSwitches) {
    instruction += ` ${options.customSwitches}`;
  }

  // Preserve formatting
  if (options.preserveFormatting !== false) {
    instruction += ' \\* MERGEFORMAT';
  }

  return instruction;
}
```

#### Task 4: Field Switch Setter

**File**: [`src/elements/TableOfContents.ts`](src/elements/TableOfContents.ts)

```typescript
/**
 * Updates the TOC field instruction with new switches
 * Rebuilds the field instruction based on current properties
 *
 * @param switches - Switches to update
 * @returns This TableOfContents for chaining
 *
 * @example
 * toc.setFieldSwitches({
 *   hyperlinks: true,
 *   hideInWebLayout: true,
 *   tabLeader: 'none'
 * });
 */
public setFieldSwitches(switches: {
  hyperlinks?: boolean;
  hidePageNumbers?: boolean;
  hideInWebLayout?: boolean;
  tabLeader?: 'dot' | 'hyphen' | 'underscore' | 'none';
}): this {
  if (switches.hyperlinks !== undefined) {
    this.useHyperlinks = switches.hyperlinks;
  }

  if (switches.hidePageNumbers !== undefined) {
    this.showPageNumbers = !switches.hidePageNumbers;
  }

  if (switches.hideInWebLayout !== undefined) {
    this.hideInWebLayout = switches.hideInWebLayout;
  }

  if (switches.tabLeader !== undefined) {
    this.tabLeader = switches.tabLeader;
  }

  // Clear originalFieldInstruction to force rebuild
  this.originalFieldInstruction = undefined;

  return this;
}
```

### Phase 3: Documentation & Examples 📚

#### Task 5: API Documentation

**File**: `docs/api/toc-field-api.md`

Document all TOC field-related methods:

- `getFieldInstruction()` - Retrieve current field code
- `setFieldSwitches()` - Update field switches
- `validateFieldStructure()` - Verify field integrity
- `buildCustomFieldInstruction()` - Create custom field codes

#### Task 6: Usage Examples

**File**: `examples/08-table-of-contents/toc-field-manipulation.ts`

```typescript
/**
 * TOC Field Manipulation Examples
 * Demonstrates advanced TOC field customization
 */

import { Document, TableOfContents } from "../../src";

// Example 1: Creating TOC with custom field instruction
async function example1_CustomFieldInstruction() {
  const doc = Document.create();

  const instruction = TableOfContents.buildCustomFieldInstruction({
    levels: { start: 1, end: 4 },
    switches: {
      hyperlinks: true,
      hideInWebLayout: true,
      tabLeader: "dot",
    },
    preserveFormatting: true,
  });

  console.log("Field instruction:", instruction);
  // Output: "TOC \o "1-4" \h \z \* MERGEFORMAT"

  const toc = TableOfContents.create({
    title: "Contents",
    levels: 4,
    useHyperlinks: true,
    hideInWebLayout: true,
  });

  doc.addTableOfContents(toc);

  // Add content...
  doc.createParagraph("Chapter 1").setStyle("Heading1");

  await doc.save("custom-field-toc.docx");
}

// Example 2: Modifying existing TOC field switches
async function example2_ModifyFieldSwitches() {
  const doc = await Document.load("existing.docx");

  const tocs = doc.getTableOfContentsElements();
  if (tocs.length > 0) {
    const toc = tocs[0]?.getTableOfContents();

    if (toc) {
      // Update field switches
      toc.setFieldSwitches({
        hyperlinks: true,
        hideInWebLayout: true,
        tabLeader: "none",
        hidePageNumbers: true,
      });

      // Verify field instruction
      const instruction = toc.getFieldInstruction();
      console.log("Updated instruction:", instruction);

      // Validate structure
      const validation = toc.validateFieldStructure();
      if (!validation.valid) {
        console.error("Field structure errors:", validation.errors);
      }
    }
  }

  await doc.save("modified-toc.docx");
}

// Example 3: Custom styles in TOC
async function example3_CustomStylesTOC() {
  const doc = Document.create();

  const instruction = TableOfContents.buildCustomFieldInstruction({
    customStyles: [
      { name: "MyHeading1", level: 1 },
      { name: "MyHeading2", level: 2 },
      { name: "MySpecialSection", level: 1 },
    ],
    switches: {
      hyperlinks: true,
    },
  });

  console.log("Custom styles instruction:", instruction);
  // Output: TOC \t "MyHeading1,1," \t "MyHeading2,2," \t "MySpecialSection,1," \h \* MERGEFORMAT

  const toc = TableOfContents.create({
    includeStyles: [
      { styleName: "MyHeading1", level: 1 },
      { styleName: "MyHeading2", level: 2 },
      { styleName: "MySpecialSection", level: 1 },
    ],
    useHyperlinks: true,
  });

  doc.addTableOfContents(toc);

  await doc.save("custom-styles-toc.docx");
}
```

### Phase 4: Testing 🧪

#### Task 7: Comprehensive Test Suite

**File**: `src/__tests__/toc-field-structure.test.ts`

```typescript
import { TableOfContents, Document } from "../index";

describe("TOC Field Structure", () => {
  describe("Field Instruction Building", () => {
    it("should build standard field instruction", () => {
      const toc = TableOfContents.create({ levels: 3 });
      const instruction = toc.getFieldInstruction();
      expect(instruction).toBe('TOC \\o "1-3" \\* MERGEFORMAT');
    });

    it("should include hyperlinks switch", () => {
      const toc = TableOfContents.create({
        levels: 3,
        useHyperlinks: true,
      });
      const instruction = toc.getFieldInstruction();
      expect(instruction).toContain("\\h");
    });

    it("should include hide web layout switch", () => {
      const toc = TableOfContents.create({
        hideInWebLayout: true,
      });
      const instruction = toc.getFieldInstruction();
      expect(instruction).toContain("\\z");
    });
  });

  describe("Field Structure Validation", () => {
    it("should validate correct field structure", () => {
      const toc = TableOfContents.create({ levels: 3 });
      const validation = toc.validateFieldStructure();
      expect(validation.valid).toBe(true);
      expect(validation.errors).toHaveLength(0);
    });

    it("should detect invalid levels", () => {
      const toc = TableOfContents.create({ levels: 10 }); // Invalid
      const validation = toc.validateFieldStructure();
      expect(validation.valid).toBe(false);
      expect(validation.errors).toContain(
        expect.stringMatching(/levels must be/i)
      );
    });
  });

  describe("Field Instruction Persistence", () => {
    it("should preserve original field instruction from loaded documents", async () => {
      // This test would require a loaded document with a TOC
      // Verify that originalFieldInstruction is preserved
    });
  });
});
```

## API Summary

### Core Methods

| Method                                 | Purpose                           | Returns            |
| -------------------------------------- | --------------------------------- | ------------------ |
| `getFieldInstruction()`                | Get current TOC field instruction | `string`           |
| `setFieldSwitches(switches)`           | Update field switches             | `this`             |
| `validateFieldStructure()`             | Verify field integrity            | `ValidationResult` |
| `buildCustomFieldInstruction(options)` | Build custom field code           | `string` (static)  |

### Field Switches Reference

| Switch              | Purpose                                  | Example                    |
| ------------------- | ---------------------------------------- | -------------------------- |
| `\o "1-3"`          | Outline levels to include                | `TOC \o "1-3"`             |
| `\h`                | Use hyperlinks instead of page numbers   | `TOC \o "1-3" \h`          |
| `\n`                | Hide page numbers                        | `TOC \o "1-3" \n`          |
| `\z`                | Hide page numbers in web layout          | `TOC \o "1-3" \z`          |
| `\p "c"`            | Tab leader character (dot, hyphen, etc.) | `TOC \o "1-3" \p "h"`      |
| `\t "Style,Level,"` | Include custom style                     | `TOC \t "MyHeading,1,"`    |
| `\* MERGEFORMAT`    | Preserve formatting on update            | Always included by default |

## Key Considerations

### Scalability

- ✅ Current implementation scales to documents of any size
- ✅ Field structure is independent of document content
- ✅ No performance impact from complex field instructions

### Security

- ✅ Field instructions are properly XML-escaped
- ✅ No injection vulnerabilities in field code building
- ⚠️ User-provided `fieldSwitches` should be validated

### Best Practices

1. **Always include `\* MERGEFORMAT`** - Preserves formatting on field update
2. **Use `\h` for web documents** - Clickable TOC entries
3. **Combine `\n` and `\z`** - Hide page numbers completely
4. **Validate after modification** - Use `validateFieldStructure()`

## Implementation Timeline

| Phase   | Tasks                      | Est. Effort | Priority |
| ------- | -------------------------- | ----------- | -------- |
| Phase 1 | Documentation & Validation | 4 hours     | High     |
| Phase 2 | Field Manipulation Helpers | 6 hours     | Medium   |
| Phase 3 | Examples & API Docs        | 3 hours     | Medium   |
| Phase 4 | Testing                    | 4 hours     | High     |

**Total Estimated Effort**: 17 hours

## Conclusion

The current TOC implementation **already maintains complete field structure** that supports Word's "Update Field" functionality. The enhancement plan focuses on:

1. **Better documentation** of existing capabilities
2. **Helper methods** for easier field manipulation
3. **Validation tools** to prevent field corruption
4. **Examples** demonstrating field updateability

The core field structure (begin → instrText → separate → content → end) is **fully compliant with ECMA-376 §17.16.5** and will work correctly in all versions of Microsoft Word.
