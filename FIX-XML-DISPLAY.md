# Fix for XML Elements Displaying as Text in Hyperlinks

## Problem

When creating hyperlinks or text runs with content that contains XML-like patterns (e.g., `<w:t xml:space="preserve">`), these patterns appear as literal text in the Word document instead of being processed as XML.

### Example of the Problem

```typescript
// This text contains XML-like content
const text = 'Important Information<w:t xml:space="preserve">1 - Not Found';
para.addHyperlink(Hyperlink.createInternal('bookmark', text));
```

**Result in Word:** The hyperlink displays as:
```
Important Information<w:t xml:space="preserve">1 - Not Found
```

## Solution

The DocXML framework now provides automatic detection and optional cleaning of XML patterns in text content.

### Features

1. **Automatic Detection**: Warns when text contains XML-like patterns
2. **Optional Cleaning**: Can automatically remove XML patterns from text
3. **Backward Compatible**: Default behavior unchanged (warnings only)

## Usage

### Option 1: Warning Only (Default)

By default, the framework warns about XML patterns but doesn't modify the text:

```typescript
// Creates a warning in console but preserves the text as-is
const hyperlink = Hyperlink.createInternal('bookmark', problematicText);
```

Console output:
```
DocXML Text Validation Warning [Hyperlink text]:
  - Text contains XML-like markup...
```

### Option 2: Automatic Cleaning

Enable the `cleanXmlFromText` option to automatically remove XML patterns:

```typescript
// For Hyperlinks
const hyperlink = new Hyperlink({
  anchor: 'bookmark',
  text: problematicText,
  formatting: {
    cleanXmlFromText: true  // ← Enable automatic cleaning
  }
});

// For Runs
const run = new Run(problematicText, {
  cleanXmlFromText: true,  // ← Enable automatic cleaning
  bold: true
});
```

**Result in Word:** Clean text without XML markup

## API Reference

### New RunFormatting Option

```typescript
interface RunFormatting {
  // ... existing options ...

  /**
   * Automatically clean XML-like patterns from text content.
   * When enabled, removes XML tags like <w:t> from text to prevent display issues.
   * Default: false (only warns about XML patterns)
   */
  cleanXmlFromText?: boolean;
}
```

### Validation Functions

New utility functions in `src/utils/validation.ts`:

```typescript
// Detect XML patterns in text
function detectXmlInText(text: string, context?: string): TextValidationResult

// Clean XML patterns from text
function cleanXmlFromText(text: string, aggressive?: boolean): string

// Main validation function used by Run and Hyperlink
function validateRunText(text: string, options?: {
  context?: string;
  autoClean?: boolean;
  aggressive?: boolean;
  warnToConsole?: boolean;
}): TextValidationResult
```

## Examples

### Example 1: Cleaning Corrupted Hyperlink Text

```typescript
const doc = Document.create();
const para = doc.createParagraph();

// Text with XML markup that would display incorrectly
const corruptedText = 'Click here<w:t>for more</w:t> information';

// Without cleaning - XML appears in document
para.addHyperlink(Hyperlink.createExternal(
  'https://example.com',
  corruptedText
));

// With cleaning - XML removed automatically
para.addHyperlink(new Hyperlink({
  url: 'https://example.com',
  text: corruptedText,
  formatting: { cleanXmlFromText: true }
}));
// Result: "Click herefor more information"
```

### Example 2: Processing Documents with XML Corruption

```typescript
async function cleanCorruptedDocument(inputPath: string, outputPath: string) {
  const doc = await Document.load(inputPath);

  // Enable cleaning for all new content
  const cleanFormatting = { cleanXmlFromText: true };

  // Process paragraphs
  doc.getParagraphs().forEach(para => {
    para.getRuns().forEach(run => {
      const text = run.getText();
      // Re-create run with cleaning enabled
      if (text.includes('<') || text.includes('&lt;')) {
        run.setText(text); // Will trigger validation
      }
    });
  });

  await doc.save(outputPath);
}
```

## Testing

Run the test scripts to verify the implementation:

```bash
# Test validation and cleaning
npx ts-node test-fix-validation.ts

# Test the specific fix
npx ts-node test-original-issue-fixed.ts
```

## Migration Guide

### For Existing Code

1. **No changes required** - Existing code continues to work with warnings
2. **To enable cleaning**, add `cleanXmlFromText: true` to formatting options
3. **Monitor console** for validation warnings during development

### Best Practices

1. **Prevention**: Avoid creating text with XML-like patterns
2. **Detection**: Watch for console warnings during development
3. **Cleaning**: Enable `cleanXmlFromText` when processing untrusted content
4. **Testing**: Always verify output in Microsoft Word

## Technical Details

### What Gets Cleaned

The cleaning function removes:
- Word XML tags like `<w:t>`, `</w:t>`, `<w:p>`, etc.
- XML attributes like `xml:space="preserve"`
- Escaped entities that become XML tags (`&lt;w:t&gt;` → `<w:t>` → removed)

### What's Preserved

- Legitimate use of angle brackets (e.g., "x < y > z")
- HTML entities in normal text
- Mathematical expressions

### Performance

- Minimal impact - validation only runs on text creation/modification
- Regex patterns optimized for common XML patterns
- Console warnings can be disabled in production

## Troubleshooting

### Issue: Warnings appear but text isn't cleaned

**Solution**: Ensure `cleanXmlFromText: true` is set in the formatting options

### Issue: Legitimate angle brackets are removed

**Solution**: Use the default mode (warnings only) or escape the brackets

### Issue: Too many console warnings

**Solution**: Disable warnings in production:
```typescript
validateRunText(text, { warnToConsole: false });
```

## Related Files

- `src/utils/validation.ts` - Validation and cleaning implementation
- `src/elements/Run.ts` - Run class with validation
- `src/elements/Hyperlink.ts` - Hyperlink class with validation
- `test-fix-validation.ts` - Test suite for the fix
- `test-original-issue-fixed.ts` - Specific test for the reported issue