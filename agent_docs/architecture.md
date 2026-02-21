# Architecture

## Module Dependency Graph

```
index.ts (public API re-exports)
  └── core/
        ├── Document.ts (main entry point, orchestrates load/save)
        │     ├── Parser.ts (DOCX → object model)
        │     ├── Generator.ts (object model → DOCX XML)
        │     └── Validator.ts (pre-save validation)
        ├── elements/ (data model: Paragraph, Run, Table, Image, Section, etc.)
        ├── formatting/ (StylesManager, NumberingManager)
        ├── managers/ (DrawingManager, ImageManager)
        └── zip/ (ZipHandler → ZipReader/ZipWriter using JSZip)
```

## Key Abstractions

- **Document**: Container for all content. Manages lifecycle (create/load/save/dispose).
- **Paragraph**: Block-level element containing Runs, Hyperlinks, Revisions, etc.
- **Run**: Inline text with character formatting (bold, italic, font, color, size).
- **Table/TableRow/TableCell**: Nested block structure with merge support.
- **Style/NumberingDefinition**: Formatting definitions referenced by paragraphs.

## Data Flow

### Load Path
1. `Document.load(buffer)` → `ZipReader.read()` extracts ZIP entries
2. `Parser.parse()` converts XML strings → object model (using XMLParser)
3. Styles, numbering, relationships are parsed into managers
4. Original XML preserved in `_original*Xml` fields for round-trip fidelity

### Save Path
1. `Document.save()` → `Generator.generate()` converts object model → XML
2. Dirty flags checked — only modified parts are regenerated
3. Unmodified parts use original XML verbatim
4. `ZipWriter.write()` assembles ZIP archive → Buffer

### Modification
- Properties set via methods (e.g., `run.setBold(true)`) update in-memory model
- No immediate XML generation — deferred to save time
- Generator checks `if (property)` before emitting — defaults naturally omitted
