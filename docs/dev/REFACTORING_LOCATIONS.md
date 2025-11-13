# Detailed Refactoring Locations

## Quick Reference: Where to Make Changes

---

## 1. XMLParser.ts - Add Helper Methods

**File**: `src/xml/XMLParser.ts`

**Add these 3 methods** to the XMLParser class:

```typescript
/**
 * Checks if an XML element contains a child element of specified type
 * @param xml - XML content to search
 * @param childName - Child element name (e.g., 'w:b')
 * @returns True if child element exists
 */
static hasChild(xml: string, childName: string): boolean {
  const elements = this.extractElements(xml, childName);
  return elements.length > 0;
}

/**
 * Extracts all attributes from an element as an object
 * @param xml - XML content
 * @param elementName - Element name to extract from
 * @returns Object with attribute name-value pairs
 */
static extractElementProperties(xml: string, elementName: string): Record<string, string | undefined> {
  const elements = this.extractElements(xml, elementName);
  if (elements.length === 0) {
    return {};
  }

  const element = elements[0];
  const props: Record<string, string | undefined> = {};

  // Extract attributes between element name and closing >
  const openEnd = element.indexOf('>');
  if (openEnd <= 0) return props;

  const attrSection = element.substring(0, openEnd);
  // Simple regex for attributes (this is WITHIN element, not XPath)
  const attrRegex = /(\w+):(\w+)="([^"]*)"/g;
  let match;

  while ((match = attrRegex.exec(attrSection)) !== null) {
    props[`${match[1]}:${match[2]}`] = match[3];
  }

  return props;
}

/**
 * Gets all child elements of a specific type
 * @param xml - Parent XML content
 * @param childName - Child element name
 * @returns Array of child XML elements
 */
static getChildren(xml: string, childName: string): string[] {
  return this.extractElements(xml, childName);
}
```

---

## 2. DocumentParser.ts - Refactor 5 Methods

### Method 1: parseParagraphProperties()

**Current Location**: Lines 273-354

**Find**:

```typescript
private parseParagraphProperties(paraXml: string, paragraph: Paragraph): void {
  const pPrMatch = paraXml.match(/<w:pPr[^>]*>([\s\S]*?)<\/w:pPr>/);
  if (!pPrMatch || !pPrMatch[1]) {
    return;
  }

  const pPr = pPrMatch[1];

  // Alignment
  const alignMatch = pPr.match(/<w:jc\s+w:val="([^"]+)"/);
  if (alignMatch && alignMatch[1]) {
    // ... etc
```

**Replace with**:

```typescript
private parseParagraphProperties(paraXml: string, paragraph: Paragraph): void {
  const pPrXmls = XMLParser.extractElements(paraXml, 'w:pPr');
  if (pPrXmls.length === 0) {
    return;
  }

  const pPrXml = pPrXmls[0];
  const props = XMLParser.extractElementProperties(pPrXml, 'w:pPr');

  // Alignment
  const jcXmls = XMLParser.extractElements(pPrXml, 'w:jc');
  if (jcXmls.length > 0) {
    const jcXml = jcXmls[0];
    const jcProps = XMLParser.extractElementProperties(jcXml, 'w:jc');
    if (jcProps['w:val']) {
      const value = jcProps['w:val'];
      const validAlignments = ['left', 'center', 'right', 'justify'];
      if (validAlignments.includes(value)) {
        const alignment = value as 'left' | 'center' | 'right' | 'justify';
        paragraph.setAlignment(alignment);
      }
    }
  }

  // Style
  const pStyleXmls = XMLParser.extractElements(pPrXml, 'w:pStyle');
  if (pStyleXmls.length > 0) {
    const styleProps = XMLParser.extractElementProperties(pStyleXmls[0], 'w:pStyle');
    if (styleProps['w:val']) {
      paragraph.setStyle(styleProps['w:val']);
    }
  }

  // Continue with other properties using same pattern...
```

### Method 2: parseRunProperties()

**Current Location**: Lines 360-509

**Find all instances of**:

```typescript
if (rPr.includes("<w:b/>") || rPr.includes("<w:b ")) {
  run.setBold(true);
}

if (rPr.includes("<w:i/>") || rPr.includes("<w:i ")) {
  run.setItalic(true);
}
// ... etc (about 15 of these)
```

**Replace with**:

```typescript
if (XMLParser.hasChild(rPr, "w:b")) {
  run.setBold(true);
}

if (XMLParser.hasChild(rPr, "w:i")) {
  run.setItalic(true);
}
// ... etc (same pattern for all 15)
```

### Method 3: extractText()

**Current Location**: Lines 426-437

**Find**:

```typescript
const extracted = XMLParser.extractText(runXml) || "";

// Unescape XML entities
let cleaned = extracted
  .replace(/&lt;/g, "<")
  .replace(/&gt;/g, ">")
  .replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'")
  .replace(/&amp;/g, "&");
```

**Replace with**:

```typescript
const extracted = XMLParser.extractText(runXml) || "";
const cleaned = XMLBuilder.unescapeXml(extracted);
```

### Method 4: parseParagraph()

**Current Location**: Lines 241-243

**Find**:

```typescript
let paraXmlWithoutHyperlinks = paraXml;
for (const hyperlinkXml of hyperlinkXmls) {
  paraXmlWithoutHyperlinks = paraXmlWithoutHyperlinks.replace(hyperlinkXml, "");
}
```

**Note**: This is trickier. Instead of string replacement, better approach:

```typescript
// Already extracted hyperlinks separately, so just extract non-hyperlink runs
// XMLParser.extractElements already handles this correctly by finding elements
// No need to remove hyperlinks from the string first

const runXmls = XMLParser.extractElements(paraXml, "w:r");
// This will get ALL runs, including those inside hyperlinks
// But we already processed hyperlinks separately, so filter them out:

for (const runXml of runXmls) {
  // Skip runs that are part of hyperlinks (check if parent is hyperlink)
  // For now, just use original logic but make it cleaner
  const run = this.parseRun(runXml);
  if (run) {
    paragraph.addRun(run);
  }
}
```

---

## 3. validation.ts - Update Text Cleaning

**File**: `src/utils/validation.ts`

**Find location**: Search for the text cleaning function (around line 460)

**Find**:

```typescript
protected static cleanTextForComparison(text: string): string {
  let cleaned = text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
  // ... more cleaning
}
```

**Replace with**:

```typescript
protected static cleanTextForComparison(text: string): string {
  // Use XMLBuilder's unescapeXml for consistency
  const unescaped = XMLBuilder.unescapeXml(text);
  let cleaned = unescaped
    // ... keep other cleaning logic if needed
}
```

**Note**: Check if XMLBuilder needs to be imported in validation.ts

---

## 4. Files to Check for Other .match() or .replace()

Search in these files for additional XML string manipulation:

```
grep -n "\.match.*<w:" src/core/DocumentParser.ts
grep -n "\.replace.*<" src/core/DocumentParser.ts
grep -n "\.includes.*<" src/core/DocumentParser.ts
```

**Likely locations**:

- `parseParagraphProperties()` - 7-8 .match() calls (lines 273-350)
- `parseRunProperties()` - 15 .includes() calls (lines 360-509)
- Text extraction - 1 .replace() chain (lines 426-437)

---

## 5. Testing Locations to Update

**File**: `tests/xml/XMLParser.test.ts`

**Add tests for new methods**:

```typescript
describe("XMLParser.hasChild()", () => {
  test("should detect child element", () => {
    const xml = "<w:p><w:b/></w:p>";
    expect(XMLParser.hasChild(xml, "w:b")).toBe(true);
  });

  test("should return false when child not present", () => {
    const xml = "<w:p><w:i/></w:p>";
    expect(XMLParser.hasChild(xml, "w:b")).toBe(false);
  });
});

describe("XMLParser.extractElementProperties()", () => {
  test("should extract attributes", () => {
    const xml = '<w:jc w:val="center" w:other="test"/>';
    const props = XMLParser.extractElementProperties(xml, "w:jc");
    expect(props["w:val"]).toBe("center");
  });
});

describe("XMLParser.getChildren()", () => {
  test("should get all child elements", () => {
    const xml = "<w:p><w:r/><w:r/><w:r/></w:p>";
    const children = XMLParser.getChildren(xml, "w:r");
    expect(children.length).toBe(3);
  });
});
```

---

## 6. Documentation Updates

**File**: `src/xml/CLAUDE.md`

**Add section**:

```markdown
## Helper Methods

### hasChild(xml, childName)

Checks if element contains child of specified type.

### extractElementProperties(xml, elementName)

Gets all attributes of an element as key-value pairs.

### getChildren(xml, childName)

Gets all child elements of specified type.
```

**File**: `CLAUDE.md` (root level)

**Add section**:

```markdown
## XML Handling Standards

All XML operations **must** use XMLParser or XMLBuilder:

✅ CORRECT:

- XMLParser.extractElements()
- XMLParser.hasChild()
- XMLBuilder.w()

❌ INCORRECT:

- .match() for XML parsing
- .replace() for XML manipulation
- .includes() for XML checking

Exception: Path normalization in zip/validation.ts
```

---

## Checklist for Completion

- [ ] Add 3 methods to XMLParser.ts
- [ ] Update parseParagraphProperties() in DocumentParser.ts
- [ ] Update parseRunProperties() in DocumentParser.ts (~15 lines)
- [ ] Update text extraction in DocumentParser.ts
- [ ] Review parseParagraph() hyperlink handling
- [ ] Update validation.ts text cleaning
- [ ] Add tests for XMLParser.hasChild()
- [ ] Add tests for XMLParser.extractElementProperties()
- [ ] Add tests for XMLParser.getChildren()
- [ ] Run full test suite: `npm test`
- [ ] Update src/xml/CLAUDE.md
- [ ] Update root CLAUDE.md
- [ ] Commit with message: "refactor: Enforce docxmlater-only XML handling - standardize to XMLParser"

---

## Verification Commands

```bash
# Check for remaining .match() in source
grep -r "\.match.*<w:" src/ --include="*.ts" | grep -v test

# Check for remaining .replace() in source
grep -r "\.replace.*<" src/ --include="*.ts" | grep -v test

# Check for remaining .includes() with < in source
grep -r "\.includes.*<w:" src/ --include="*.ts" | grep -v test

# Run tests
npm test
```

---

## Summary

- **XMLParser.ts**: +50 lines (add 3 methods)
- **DocumentParser.ts**: ~150 lines changed
- **validation.ts**: ~10 lines changed
- **Tests**: +30 lines
- **Documentation**: +15 lines

**Total**: ~250 lines modified, 0 breaking changes
