# parseToObject() Implementation Plan

**Session Date**: October 18, 2025
**Feature**: XML-to-Object parsing for Office Open XML (OOXML) documents
**Target Class**: `XMLParser` (src/xml/XMLParser.ts)

## Problem Analysis

### Current State

- XMLParser has position-based string extraction methods
- No structured object representation of XML
- Manual attribute/element parsing using regex/indexOf
- No standard way to traverse XML as JavaScript objects

### Requirements

Implement `parseToObject()` method that converts XML strings to JavaScript objects following fast-xml-parser conventions:

- Attributes → `@_` prefix (e.g., `@_Id`, `@_Type`)
- Text content → `#text` property
- Multiple child elements → Array `[]`
- Single child element → Object `{}`
- Namespaces → Preserved in keys (e.g., `w:p`, `w:r`)
- Self-closing tags → Empty object `{}`

## Solution Architecture

### Implementation Strategy

**Parser Design:**

- Recursive descent parser for nested XML
- Stack-based approach for tracking element hierarchy
- Attribute extraction and prefixing
- Text node handling with whitespace trimming
- Array coalescing for duplicate element names

### Phase 1: Type Definitions (Priority: HIGH)

**File**: `src/xml/XMLParser.ts`

**Add Types:**

```typescript
/**
 * Options for XML-to-object parsing
 */
export interface ParseToObjectOptions {
  /** Ignore attributes (default: false) */
  ignoreAttributes?: boolean;

  /** Attribute name prefix (default: '@_') */
  attributeNamePrefix?: string;

  /** Text node property name (default: '#text') */
  textNodeName?: string;

  /** Remove namespace prefixes from element names (default: false) */
  ignoreNamespace?: boolean;

  /** Parse numeric attribute values (default: true) */
  parseAttributeValue?: boolean;

  /** Trim whitespace from text values (default: true) */
  trimValues?: boolean;

  /** Always return arrays for elements (default: false) */
  alwaysArray?: boolean;
}

/**
 * Parsed XML object structure
 * Can be a string, object, array, or nested structure
 */
export type ParsedXMLValue =
  | string
  | number
  | boolean
  | ParsedXMLObject
  | ParsedXMLObject[]
  | null
  | undefined;

/**
 * Parsed XML object with dynamic keys
 */
export interface ParsedXMLObject {
  [key: string]: ParsedXMLValue;
}
```

### Phase 2: Core Parser Implementation (Priority: HIGH)

**Method Signature:**

```typescript
static parseToObject(
  xml: string,
  options?: ParseToObjectOptions
): ParsedXMLObject
```

**Algorithm Steps:**

1. Validate and sanitize input XML
2. Remove XML declaration if present
3. Parse root element and recursively process children
4. Extract attributes and prefix with `@_`
5. Handle text content with `#text` property
6. Coalesce duplicate child elements into arrays
7. Return structured object

**Implementation Details:**

```typescript
static parseToObject(
  xml: string,
  options?: ParseToObjectOptions
): ParsedXMLObject {
  // Default options
  const opts: Required<ParseToObjectOptions> = {
    ignoreAttributes: options?.ignoreAttributes ?? false,
    attributeNamePrefix: options?.attributeNamePrefix ?? '@_',
    textNodeName: options?.textNodeName ?? '#text',
    ignoreNamespace: options?.ignoreNamespace ?? false,
    parseAttributeValue: options?.parseAttributeValue ?? true,
    trimValues: options?.trimValues ?? true,
    alwaysArray: options?.alwaysArray ?? false,
  };

  // Validate input
  XMLParser.validateSize(xml);

  // Remove XML declaration
  xml = xml.replace(/<\?xml[^>]*\?>\s*/g, '');

  // Parse recursively
  return XMLParser.parseElement(xml, 0, opts).value;
}
```

### Phase 3: Helper Methods (Priority: HIGH)

**parseElement()** - Recursive element parser:

```typescript
private static parseElement(
  xml: string,
  startPos: number,
  options: Required<ParseToObjectOptions>
): { value: ParsedXMLObject; endPos: number }
```

**extractElementName()** - Extract tag name:

```typescript
private static extractElementName(xml: string, startPos: number): string
```

**extractAttributes()** - Parse all attributes:

```typescript
private static extractAttributes(
  xml: string,
  startPos: number,
  endPos: number,
  options: Required<ParseToObjectOptions>
): Record<string, string | number | boolean>
```

**parseAttributeValue()** - Convert attribute strings to proper types:

```typescript
private static parseAttributeValue(value: string): string | number | boolean
```

**coalesceChildren()** - Merge duplicate elements into arrays:

```typescript
private static coalesceChildren(
  children: Array<{ name: string; value: ParsedXMLValue }>,
  options: Required<ParseToObjectOptions>
): ParsedXMLObject
```

### Phase 4: Testing (Priority: HIGH)

**Test File**: `tests/xml/XMLParser-parseToObject.test.ts`

**Test Coverage:**

1. **Basic Parsing**

   - Single element with attributes
   - Nested elements
   - Self-closing tags
   - Text content

2. **Relationships XML** (Real-world use case)

   - Multiple Relationship elements → Array
   - Single Relationship element → Object
   - Attribute extraction (@\_Id, @\_Type, etc.)

3. **Styles XML** (Complex nested)

   - Namespace preservation (w:style, w:pPr)
   - Nested properties
   - Mixed attributes and children

4. **Document XML** (Text content)

   - Text nodes with #text property
   - Empty elements
   - Whitespace handling

5. **Edge Cases**

   - Empty XML
   - Malformed XML handling
   - Large documents (size validation)
   - Special characters in text/attributes
   - Namespace handling

6. **Options Testing**
   - ignoreAttributes
   - attributeNamePrefix customization
   - textNodeName customization
   - ignoreNamespace
   - parseAttributeValue
   - trimValues

**Example Tests:**

```typescript
describe("XMLParser.parseToObject", () => {
  it("should parse single Relationship element", () => {
    const xml = `
      <?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://example.com/type" Target="https://example.com"/>
      </Relationships>
    `;

    const result = XMLParser.parseToObject(xml);

    expect(result).toEqual({
      Relationships: {
        "@_xmlns":
          "http://schemas.openxmlformats.org/package/2006/relationships",
        Relationship: {
          "@_Id": "rId1",
          "@_Type": "http://example.com/type",
          "@_Target": "https://example.com",
        },
      },
    });
  });

  it("should parse multiple Relationship elements as array", () => {
    const xml = `
      <Relationships>
        <Relationship Id="rId1" Target="https://example.com"/>
        <Relationship Id="rId2" Target="https://google.com"/>
      </Relationships>
    `;

    const result = XMLParser.parseToObject(xml);

    expect(result.Relationships.Relationship).toBeInstanceOf(Array);
    expect(result.Relationships.Relationship).toHaveLength(2);
  });
});
```

### Phase 5: Integration & Documentation (Priority: MEDIUM)

**Update exports** in `src/index.ts`:

```typescript
export {
  XMLParser,
  ParseToObjectOptions,
  ParsedXMLValue,
  ParsedXMLObject,
} from "./xml/XMLParser";
```

**Add JSDoc documentation**:

- Comprehensive examples in method comments
- Link to OOXML parsing patterns
- Reference fast-xml-parser compatibility

**Update CLAUDE.md**:

- Document parseToObject() usage
- Add examples for common OOXML patterns
- Note compatibility with fast-xml-parser

## Implementation Checklist

### Core Implementation

- [ ] Add type definitions (ParseToObjectOptions, ParsedXMLObject, etc.)
- [ ] Implement parseToObject() main method
- [ ] Implement parseElement() recursive parser
- [ ] Implement extractElementName() helper
- [ ] Implement extractAttributes() helper
- [ ] Implement parseAttributeValue() type conversion
- [ ] Implement coalesceChildren() array handling
- [ ] Add error handling for malformed XML

### Testing

- [ ] Write basic parsing tests (10 tests)
- [ ] Write Relationships XML tests (5 tests)
- [ ] Write Styles XML tests (5 tests)
- [ ] Write Document XML tests (5 tests)
- [ ] Write edge case tests (8 tests)
- [ ] Write options configuration tests (6 tests)
- [ ] Verify all tests pass (39 total tests)

### Documentation

- [ ] Add comprehensive JSDoc to parseToObject()
- [ ] Document all helper methods
- [ ] Add code examples to comments
- [ ] Update src/index.ts exports
- [ ] Update CLAUDE.md with parseToObject() section

### Integration

- [ ] Ensure no breaking changes to existing XMLParser methods
- [ ] Verify compatibility with DocumentParser usage
- [ ] Test with real DOCX files
- [ ] Performance test with large XML documents

## Affected Files

**Modified Files:**

- `src/xml/XMLParser.ts` (add parseToObject() + helpers, ~250 lines)
- `src/index.ts` (export new types)
- `CLAUDE.md` (document new feature)

**New Files:**

- `tests/xml/XMLParser-parseToObject.test.ts` (NEW, ~400 lines)

## Success Criteria

✅ parseToObject() returns objects matching fast-xml-parser format
✅ Attributes use @\_ prefix
✅ Text content uses #text property
✅ Single elements → Object, multiple elements → Array
✅ Namespaces preserved in element names
✅ Self-closing tags → Empty objects
✅ All 39 tests pass
✅ No breaking changes to existing code
✅ Performance acceptable (< 100ms for typical DOCX parts)
✅ Comprehensive documentation and examples

## Risk Mitigation

**Risk**: Malformed XML could cause infinite loops
**Mitigation**: Add position tracking, max iteration limits, validateSize()

**Risk**: Large XML documents could cause memory issues
**Mitigation**: Use existing validateSize() check (10MB limit)

**Risk**: Edge cases in attribute/text parsing
**Mitigation**: Comprehensive test suite covering all OOXML patterns

**Risk**: Breaking changes to existing XMLParser users
**Mitigation**: New method is additive, all existing methods unchanged

## Timeline Estimate

- Phase 1: Type definitions (10 min)
- Phase 2: Core parseToObject() implementation (45 min)
- Phase 3: Helper methods (30 min)
- Phase 4: Testing (40 min)
- Phase 5: Documentation (15 min)
- **Total**: ~140 min (~2.5 hours)

## Version Impact

- **Current**: 0.10.0 (current version)
- **Next**: 0.11.0 (XML parsing enhancement)
- **Breaking Changes**: None
- **New APIs**: parseToObject(), ParseToObjectOptions, ParsedXMLObject
- **Deprecations**: None

## Implementation Notes

### Parsing Strategy

**Position-Based Parsing:**
The implementation will use position-based parsing (consistent with existing XMLParser methods) rather than regex to avoid ReDoS vulnerabilities.

**Algorithm Flow:**

1. Find opening tag `<elementName`
2. Extract element name (handle namespaces)
3. Extract attributes until `>`
4. Check for self-closing `/>` or full closing `</elementName>`
5. If has children, recursively parse
6. If has text, extract and add as #text
7. Coalesce duplicate child element names into arrays

**Example Walkthrough:**

```xml
<Relationships>
  <Relationship Id="rId1"/>
  <Relationship Id="rId2"/>
</Relationships>
```

Parse steps:

1. Parse `<Relationships>` → Create object `{ Relationships: {} }`
2. Parse first `<Relationship>` → Add `{ '@_Id': 'rId1' }`
3. Parse second `<Relationship>` → Detect duplicate name
4. Coalesce into array: `{ Relationship: [{ '@_Id': 'rId1' }, { '@_Id': 'rId2' }] }`

This ensures single vs. multiple element handling matches fast-xml-parser behavior.
