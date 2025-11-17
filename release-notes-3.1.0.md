## 🎉 Version 3.1.0 - TOC Improvements

### ✨ Added

- **TOC Range Format Support**: Enhanced `\t` switch to support numeric range format
  - New range format: `\t "2-3"` similar to `\o` switch behavior
  - Supports patterns like `\t "2-2"` → [2], `\t "2-3"` → [2, 3], `\t "1-5"` → [1, 2, 3, 4, 5]
  - Maintains backward compatibility with style name format: `\t "Heading 2,2,"`
  - Parser detects range format via regex `/^(\d+)-(\d+)$/` before processing style names

### 🐛 Fixed

- **TOC Field Instruction Parsing**: Fixed critical bug where TOCs with ONLY `\t` switches incorrectly fell back to default levels [1,2,3]
  - **Root cause**: `parseTOCFieldInstruction()` returned default [1,2,3] whenever `levels.size === 0`, regardless of whether switches were present
  - **Issue**: Field instruction `TOC \h \u \z \t "Heading 2,2,"` should ONLY include Heading 2 paragraphs, but incorrectly included Heading 1 as well
  - **Solution**: Track whether `\t`, `\o`, or `\u` switches were found during parsing
  - Now only uses default [1,2,3] when NO switches are present
  - Returns empty array when switches exist but resulted in empty levels
  - Added support for `\u` switch (use outline levels from paragraph formatting)

### 📝 Examples

```
TOC \t "Heading 2,2,"           → [2] (not [1,2,3])
TOC \h \u \z \t "Heading 2,2,"  → [2] (not [1,2,3])
TOC \o "1-3"                    → [1,2,3]
TOC                              → [1,2,3] (default when no switches)
TOC \t "2-3"                    → [2, 3] (new range format)
```

### 📦 Installation

```bash
npm install docxmlater@3.1.0
```

---

**Full Changelog**: https://github.com/ItMeDiaTech/docXMLater/compare/v3.0.0...v3.1.0
