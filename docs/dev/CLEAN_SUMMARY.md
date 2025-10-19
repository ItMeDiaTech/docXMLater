# Clean Summary: Enforce docxmlater Framework

---

## ✅ What's Already Correct

1. **JSZip** - Only used for ZIP operations (GOOD - keep it)
2. **XMLParser** - Your framework, position-based parsing (GOOD - it's there)
3. **XMLBuilder** - Your framework, for generation (GOOD - it's there)

---

## ⚠️ What Needs Fixing

**File**: `src/core/DocumentParser.ts`

Uses **direct regex** (`.match()`, `.includes()`, `.replace()`) instead of **XMLParser methods** that already exist.

---

## The Fix: ~50 Lines in One File

Replace these patterns in DocumentParser.ts:

| Find                                              | Replace With                                              |
| ------------------------------------------------- | --------------------------------------------------------- |
| `.match(/<w:pPr[^>]*>(.*)<\/w:pPr>/)`             | `XMLParser.extractBetweenTags(xml, '<w:pPr', '</w:pPr>')` |
| `.match(/w:val="([^"]*)"/)`                       | `XMLParser.extractAttribute(xml, 'w:val')`                |
| `.includes('<w:tag')`                             | `XMLParser.hasSelfClosingTag(xml, 'w:tag')`               |
| `.replace(/&lt;/g, '<').replace(/&gt;/g, '>')...` | `XMLBuilder.unescapeXml(text)`                            |

---

## Specific Locations

**DocumentParser.ts**:

1. **Lines 274-279** - Replace `.match()` for `<w:pPr>`
2. **Lines 282, 295, 301, 320** - Replace `.match()` for attributes
3. **Line 304-306** - Replace `.match()` for attribute values
4. **Lines 349-351** - Replace `.includes()` checks
5. **Lines 501-507** - Replace `.includes()` checks in parseRunProperties
6. **Lines 426-431** - Replace `.replace()` chain with `XMLBuilder.unescapeXml()`

---

## Result

**Before**: Mixed regex + XMLParser

```typescript
const pPrMatch = paraXml.match(/<w:pPr[^>]*>([\s\S]*?)<\/w:pPr>/); // ⚠️ Regex bypass
if (!pPrMatch) return;
const pPr = pPrMatch[1];
```

**After**: 100% XMLParser

```typescript
const pPr = XMLParser.extractBetweenTags(paraXml, "<w:pPr", "</w:pPr>"); // ✅ Using framework
if (!pPr) return;
```

---

## Checklist

- [ ] Review `CORRECTED_AUDIT.md` for all locations
- [ ] Replace regex patterns with XMLParser methods
- [ ] Replace entity unescaping with XMLBuilder.unescapeXml()
- [ ] Run: `npm test`
- [ ] Verify no regex `.match()` remains for XML
- [ ] Commit: "refactor: Use XMLParser exclusively in DocumentParser"

---

## No Breaking Changes

- All changes are internal to DocumentParser
- Public API stays the same
- All tests should pass
- Same functionality, consistent framework usage
