# CORRECTED Framework Audit

## Replace Direct Regex with XMLParser Methods

---

## What I Found (Corrected)

### XMLParser Already Has These Methods ✅

Your XMLParser **already provides** everything needed:

```typescript
// In src/xml/XMLParser.ts (already exist!)
static extractAttribute(xml: string, attributeName: string): string | undefined
static hasSelfClosingTag(xml: string, tagName: string): boolean
static extractBetweenTags(xml: string, startTag: string, endTag: string): string | undefined
static extractElements(xml: string, tagName: string): string[]
static extractText(xml: string): string
```

### The Real Problem: DocumentParser.ts Uses Regex Instead

**File**: `src/core/DocumentParser.ts`

Uses `.match()` when `XMLParser.extractBetweenTags()` already exists:

**Line 274 (WRONG)**:

```typescript
const pPrMatch = paraXml.match(/<w:pPr[^>]*>([\s\S]*?)<\/w:pPr>/);
const pPr = pPrMatch[1]; // Extract the inner content
```

**Should be (CORRECT)**:

```typescript
const pPr = XMLParser.extractBetweenTags(paraXml, "<w:pPr", "</w:pPr>");
if (!pPr) return;
```

---

## All Instances to Fix

### Problem 1: `.match()` for Element Extraction

Uses regex when `XMLParser.extractBetweenTags()` exists

**Lines in DocumentParser.ts**:

- Line 274: `pPrMatch = paraXml.match(/<w:pPr[^>]*>([\s\S]*?)<\/w:pPr>/);`
- Line 282: `alignMatch = pPr.match(/<w:jc\s+w:val="([^"]+)"/);`
- Line 295: `styleMatch = pPr.match(/<w:pStyle\s+w:val="([^"]+)"/);`
- Line 301: `indMatch = pPr.match(/<w:ind([^>]+)\/>/);`
- Line 320: `spacingMatch = pPr.match(/<w:spacing([^>]+)\/>/);`

**Should use**: `XMLParser.extractBetweenTags()`

### Problem 2: `.includes()` for Element Detection

Uses string search when `XMLParser.hasSelfClosingTag()` exists

**Lines in DocumentParser.ts**:

- Line 349: `if (pPr.includes('<w:keepNext')`
- Line 350: `if (pPr.includes('<w:keepLines')`
- Line 351: `if (pPr.includes('<w:pageBreakBefore')`
- Lines 501-507: Multiple `.includes()` checks in `parseRunProperties()`

**Should use**: `XMLParser.hasSelfClosingTag()`

### Problem 3: `.match()` for Attribute Extraction

Uses regex when `XMLParser.extractAttribute()` exists

**Lines in DocumentParser.ts**:

- Line 304: `leftMatch = indStr.match(/w:left="(\d+)"/);`
- Line 305: `rightMatch = indStr.match(/w:right="(\d+)"/);`
- Line 306: `firstLineMatch = indStr.match(/w:firstLine="(\d+)"/);`
- Lines 323-325: Similar attribute extraction patterns

**Should use**: `XMLParser.extractAttribute()`

### Problem 4: `.replace()` for String Manipulation

Uses string replacement when `XMLBuilder.unescapeXml()` exists

**Line 243**:

```typescript
paraXmlWithoutHyperlinks = paraXmlWithoutHyperlinks.replace(hyperlinkXml, "");
```

**Lines 426-431**: XML entity unescaping with `.replace()`

**Should use**: `XMLBuilder.unescapeXml()` for entities

---

## Fix Summary

| Pattern                    | Current                           | Should Use                       | Lines            |
| -------------------------- | --------------------------------- | -------------------------------- | ---------------- |
| Extract element content    | `.match(/<tag[^>]*>(.*)<\/tag>/)` | `XMLParser.extractBetweenTags()` | 274, 301, 320    |
| Extract attribute value    | `.match(/attr="([^"]*)"/)`        | `XMLParser.extractAttribute()`   | 304-306, 323-325 |
| Check for element          | `.includes('<w:tag')`             | `XMLParser.hasSelfClosingTag()`  | 349-351, 501-507 |
| Unescape XML entities      | `.replace(/&lt;/g, '<')`          | `XMLBuilder.unescapeXml()`       | 426-431          |
| Remove element from string | `.replace(element, '')`           | Refactor logic (don't remove)    | 243              |

---

## Corrected Refactoring Guide

### Example 1: Parse Paragraph Properties

**CURRENT (Wrong)**:

```typescript
const pPrMatch = paraXml.match(/<w:pPr[^>]*>([\s\S]*?)<\/w:pPr>/);
if (!pPrMatch || !pPrMatch[1]) {
  return;
}
const pPr = pPrMatch[1];
```

**CORRECTED (Right)**:

```typescript
const pPr = XMLParser.extractBetweenTags(paraXml, "<w:pPr", "</w:pPr>");
if (!pPr) {
  return;
}
```

### Example 2: Extract Attribute Value

**CURRENT (Wrong)**:

```typescript
const alignMatch = pPr.match(/<w:jc\s+w:val="([^"]+)"/);
if (alignMatch && alignMatch[1]) {
  const value = alignMatch[1];
  // ...
}
```

**CORRECTED (Right)**:

```typescript
// Get the jc element
const jcXml = XMLParser.extractElements(pPr, "w:jc")[0];
if (jcXml) {
  const value = XMLParser.extractAttribute(jcXml, "w:val");
  if (value) {
    // ...
  }
}
```

### Example 3: Check for Element Presence

**CURRENT (Wrong)**:

```typescript
if (pPr.includes("<w:keepNext")) {
  paragraph.setKeepNext(true);
}
```

**CORRECTED (Right)**:

```typescript
if (XMLParser.hasSelfClosingTag(pPr, "w:keepNext")) {
  paragraph.setKeepNext(true);
}
```

### Example 4: Unescape XML

**CURRENT (Wrong)**:

```typescript
let cleaned = extracted
  .replace(/&lt;/g, "<")
  .replace(/&gt;/g, ">")
  .replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'")
  .replace(/&amp;/g, "&");
```

**CORRECTED (Right)**:

```typescript
const cleaned = XMLBuilder.unescapeXml(extracted);
```

---

## Files to Modify

1. **`src/core/DocumentParser.ts`** - ~50 lines to change

   - parseParagraphProperties() method
   - parseRunProperties() method
   - Text extraction logic

2. **`src/utils/validation.ts`** - ~5 lines
   - Text entity unescaping

---

## No New Code Needed

✅ **XMLParser already has all required methods**
✅ **XMLBuilder already has unescapeXml()**
✅ **Just need to use them instead of regex**

---

## Testing

After refactoring, verify:

- `npm test` - All existing tests pass
- No regex pattern bypassing XMLParser
- All XML operations use framework methods

---

## Verification Command

```bash
# Find all remaining .match() calls with XML patterns
grep -n "\.match.*<w:" src/core/DocumentParser.ts

# Find all remaining .includes() checks
grep -n "\.includes.*<w:" src/core/DocumentParser.ts

# Find all remaining .replace() for XML
grep -n "\.replace.*[&<>]" src/core/DocumentParser.ts
```

Expected: All should return 0 matches after refactoring.
