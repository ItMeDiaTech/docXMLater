# docXMLater Feature Showcase

## Overview

The `showcase.docx` file is a comprehensive demonstration document that showcases all features currently available in the docXMLater library. This document was generated entirely programmatically using the library's API.

## What's Inside

The showcase document demonstrates the following features:

### 1. Text Formatting
- **Character formatting**: bold, italic, underline, strikethrough
- **Subscript and superscript**
- **Font variations**: Arial, Times New Roman, Courier New
- **Size variations**: 8pt, 12pt, 18pt, 24pt
- **Colors**: red, green, blue, purple
- **Highlighting**: yellow, cyan, lightGray
- **Text effects**: small caps, all caps
- **Combinations**: multiple formats applied simultaneously

### 2. Paragraph Formatting
- **Alignment**: left, center, right, justify
- **Indentation**: first-line, hanging, left, right, both
- **Spacing**: before, after, line spacing (double spacing example)

### 3. Numbered Lists
- Multi-level numbered lists (up to 3 levels shown)
- Automatic numbering with proper hierarchical structure
- Demonstrates list continuation and sub-items

### 4. Bulleted Lists
- Multi-level bulleted lists (up to 3 levels shown)
- Hierarchical bullet structure
- Mixed-level list items

### 5. Tables
- **Basic table**: Simple 4x3 table with header row
- **Advanced table**: Complex table with:
  - Cell spanning (merged cells)
  - Custom borders (colored, double-line)
  - Cell shading (header and alternating rows)
  - Column width control
  - Cell alignment (center, left)
  - Bold and colored text in cells

### 6. Custom Styles
- **Showcase Title**: Large centered blue title
- **Showcase Subtitle**: Centered italic subtitle
- **Showcase Body**: Justified text with first-line indent
- **Showcase Code**: Monospace code style
- **Showcase Quote**: Indented quote style
- **Heading 1, 2, 3**: Standard heading hierarchy

### 7. Advanced Features
Documented features include:
- ZIP archive handling
- UTF-8 encoding support
- XML generation (ECMA-376 compliant)
- Style inheritance and cascading
- Multi-level numbering
- Table cell spanning and merging
- Section configuration

### 8. Document Statistics
A table showing library metrics:
- Library Version: v0.27.0
- Total Test Suite: 226+ tests
- Source Files: 48 files
- Lines of Code: ~10,000+
- Test Coverage: >90%

## Running the Showcase

To regenerate the showcase document:

```bash
npx ts-node showcase.ts
```

This will create a new `showcase.docx` file in the project root directory.

## File Structure

The showcase document contains:
- 89 paragraphs
- 3 tables
- 9 custom styles
- ~700 words
- File size: ~8.3 KB

## Purpose

This showcase serves multiple purposes:

1. **Feature Demonstration**: Shows what the library can do
2. **API Examples**: Demonstrates how to use various API methods
3. **Testing**: Validates that all features work together
4. **Documentation**: Provides a visual reference for capabilities
5. **Quality Assurance**: Ensures generated documents open correctly in Microsoft Word

## Compatibility

The generated document is:
- Compliant with ECMA-376 (Office Open XML) standards
- Compatible with Microsoft Word 2016 and later
- Opens correctly in LibreOffice Writer
- Readable by other DOCX-compatible applications

## Source Code

The source code for generating this showcase is in [`showcase.ts`](./showcase.ts). Review this file to see practical examples of the docXMLater API in action.
