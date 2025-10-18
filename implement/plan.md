# Font Embedding Implementation Plan

**Session Date**: October 18, 2025
**Issue**: DOCX document references embedded fonts (e.g., Play-regular.ttf) but fonts are not properly registered in [Content_Types].xml, causing OOXML validation errors.

## Problem Analysis

### Current State

- Document can reference font files but no mechanism to embed/register them
- [Content_Types].xml doesn't include font file MIME types
- word/\_rels/document.xml.rels doesn't link to font resources
- No FontManager or helper functions for font operations
- Font files (.ttf, .otf) not handled in ZIP archive operations

### Error Observed

```
Found 1 OOXML issues: Document references font Play-regular.ttf but not properly registered
```

### Root Cause

When a document references an embedded font, OOXML requires:

1. Font file added to `word/fonts/` directory in the archive
2. MIME type entry in `[Content_Types].xml` for font file format
3. Relationship entry in `word/_rels/document.xml.rels` linking to font
4. Font reference in styles or document XML pointing to embedded font

Currently: Only #4 exists (font reference), #1-3 are missing

## Solution Architecture

### Phase 1: Core Font Infrastructure (Priority: HIGH)

Create `src/elements/FontManager.ts` - Manages embedded fonts similar to ImageManager

**Responsibilities:**

- Register embedded fonts from file paths or buffers
- Track font relationships (ID, filename, MIME type)
- Generate relationship IDs (`rId1`, `rId2`, etc.)
- Validate font formats (.ttf, .otf)
- Manage font metadata (name, format, size)

**Key Methods:**

```typescript
registerFont(fontPath: string, fontName?: string): FontEntry
registerFontFromBuffer(buffer: Buffer, filename: string, fontName?: string): FontEntry
getFont(fontId: string): FontEntry | undefined
getAllFonts(): FontEntry[]
getFontCount(): number
clear(): void
```

### Phase 2: Content_Types Registration (Priority: HIGH)

Update `src/core/DocumentGenerator.ts` - Add font MIME types

**Changes:**

- Add font extension handlers (.ttf → `application/x-font-ttf`)
- Add font extension handlers (.otf → `application/x-font-opentype`)
- Update `generateContentTypesWithImagesHeadersFootersAndComments()` to include fonts

**Font MIME Types:**

- `.ttf`: `application/x-font-ttf` (TrueType font)
- `.otf`: `application/x-font-opentype` (OpenType font)
- `.woff`: `application/font-woff` (Web Open Font Format)
- `.woff2`: `application/font-woff2` (WOFF2 format)

### Phase 3: Relationship Management (Priority: HIGH)

Update `src/core/RelationshipManager.ts` - Add font relationship methods

**New Methods:**

```typescript
addFont(target: string): Relationship
addFontRelationship(relId: string, target: string): void
```

**Relationship Type:**

- Type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable`

### Phase 4: Document Integration (Priority: MEDIUM)

Update `src/core/Document.ts` - Add font management API

**New Properties:**

- `fontManager: FontManager` - Similar to imageManager

**New Methods:**

```typescript
embedFont(fontPath: string, fontName?: string): Font
embedFontFromBuffer(buffer: Buffer, filename: string): Font
getFont(fontId: string): Font | undefined
getAllFonts(): Font[]
```

### Phase 5: ZIP Archive Handling (Priority: MEDIUM)

Update `src/zip/ZipWriter.ts` and `ZipHandler.ts`

**No changes needed** - Already handles binary files correctly via Buffer operations

### Phase 6: Helper Functions (Priority: MEDIUM)

Create utility functions for font operations

**Location**: `src/utils/fontUtils.ts`

**Functions:**

```typescript
validateFontFile(buffer: Buffer, filename: string): { valid: boolean; format: string }
getFontMimeType(filename: string): string
generateFontFilename(originalName: string, index: number): string
parseFontName(fontPath: string): string
```

### Phase 7: Testing (Priority: HIGH)

Create comprehensive test suite

**Test File**: `tests/elements/FontManager.test.ts`

**Test Cases:**

- [ ] Register font from file path
- [ ] Register font from buffer
- [ ] Validate font file formats
- [ ] Font relationship generation
- [ ] Content_Types registration
- [ ] Round-trip: embed → save → load → verify
- [ ] Multiple font registration
- [ ] Font name handling
- [ ] Binary font file preservation

## Implementation Checklist

### Core Implementation

- [ ] Create FontEntry interface and types
- [ ] Create FontManager class (80 lines)
- [ ] Update DocumentGenerator with font MIME types (50 lines)
- [ ] Update RelationshipManager with font methods (30 lines)
- [ ] Update Document class with font API (40 lines)
- [ ] Create fontUtils helper functions (60 lines)

### Integration

- [ ] Update Document.save() to save font files to ZIP
- [ ] Update [Content_Types].xml generation for fonts
- [ ] Update document.\_rels.xml.rels for font relationships
- [ ] Ensure font files saved to word/fonts/ directory

### Testing

- [ ] Write 8-10 font embedding tests
- [ ] Test font file preservation through ZIP roundtrip
- [ ] Test OOXML compliance (no validation errors)
- [ ] Test multiple fonts in single document
- [ ] Test font formats (.ttf, .otf, etc.)

### Documentation

- [ ] Add JSDoc to FontManager
- [ ] Create fontUtils documentation
- [ ] Update CLAUDE.md with font embedding section
- [ ] Add font embedding example

## Affected Files

**New Files:**

- `src/elements/FontManager.ts` (NEW)
- `src/utils/fontUtils.ts` (NEW)
- `tests/elements/FontManager.test.ts` (NEW)
- `examples/fonts/embed-fonts.ts` (NEW)

**Modified Files:**

- `src/core/Document.ts` (add font manager integration)
- `src/core/DocumentGenerator.ts` (add font MIME types)
- `src/core/RelationshipManager.ts` (add font relationship methods)
- `src/zip/types.ts` (add font paths)
- `CLAUDE.md` (document font embedding)

## Success Criteria

✅ Embedded fonts appear in [Content_Types].xml with correct MIME type
✅ Font files saved to word/fonts/ directory in DOCX archive
✅ Font relationships appear in word/\_rels/document.xml.rels
✅ Fonts preserved through save/load roundtrip
✅ OOXML validation passes (no corruption warnings)
✅ All 8-10 font tests pass
✅ Supports .ttf, .otf, .woff, .woff2 formats
✅ FontManager follows ImageManager design pattern

## Risk Mitigation

**Risk**: Font files could bloat DOCX file size
**Mitigation**: Add optional font subsetting (future phase), document file size impact

**Risk**: Unsupported font formats could cause issues
**Mitigation**: Validate format on registration, throw clear error

**Risk**: Duplicate font registration
**Mitigation**: Use filename as unique key, warn on duplicate names

## Timeline Estimate

- Phase 1-2: 30 min (FontManager + MIME types)
- Phase 3-4: 20 min (RelationshipManager + Document integration)
- Phase 5-6: 15 min (ZIP handling review + utilities)
- Phase 7: 30 min (Testing + validation)
- Total: ~95 min for complete implementation

## Version Impact

- **Current**: 0.8.0 (UTF-8 encoding)
- **Next**: 0.9.0 (Font Embedding)
- **Breaking Changes**: None
- **New APIs**: FontManager, embedFont(), embedFontFromBuffer()
- **Deprecations**: None
