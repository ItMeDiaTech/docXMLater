# OOXML Glossary

Key ECMA-376 terminology used throughout this codebase.

## Package Structure

- **OPC** (Open Packaging Conventions): ZIP-based container format. A .docx file is an OPC package.
- **Part**: A file inside the ZIP (e.g., `word/document.xml`). Each part has a content type.
- **Content Type**: MIME type for a part, declared in `[Content_Types].xml`.
- **Relationship**: A link between parts. Stored in `_rels/` directories. Each has an rId, type, and target.

## WML (WordprocessingML)

- **w:document**: Root element of `word/document.xml`.
- **w:body**: Container for all block-level content.
- **w:p**: Paragraph element.
- **w:pPr**: Paragraph properties (alignment, spacing, indentation, style, numbering).
- **w:r**: Run element — inline content with uniform formatting.
- **w:rPr**: Run properties (bold, italic, font, color, size, underline).
- **w:t**: Text element inside a run. `xml:space="preserve"` to keep whitespace.
- **w:tbl / w:tr / w:tc**: Table, table row, table cell.
- **w:sectPr**: Section properties (page size, margins, orientation, headers/footers).

## Formatting

- **w:pStyle / w:rStyle**: Reference to a style definition by styleId.
- **w:numPr**: Numbering properties — links paragraph to a numbering definition.
- **w:ind**: Indentation (left, right, hanging, firstLine) in twips.
- **w:spacing**: Paragraph spacing (before, after, line) in twips/240ths.
- **w:jc**: Justification (left, center, right, both).

## Identifiers

- **rId**: Relationship ID — links elements to parts (e.g., images, hyperlinks).
- **rsid**: Revision Save ID — Word assigns unique IDs per editing session. OPTIONAL per spec.
- **numId / abstractNumId**: Numbering definition identifiers for lists.
- **cnfStyle**: Conditional formatting bitmask for table style conditionals.

## Units

- **Twips**: 1/20 of a point. 1 inch = 1440 twips.
- **EMUs**: English Metric Units. 1 inch = 914400 EMUs. Used for drawings.
- **Half-points**: Used for font sizes. 1 point = 2 half-points.

## Namespaces

| Prefix | URI | Usage |
|--------|-----|-------|
| w | wordprocessingml/2006/main | Core WML elements |
| r | officeDocument/2006/relationships | Relationship references |
| wp | drawingml/2006/wordprocessingDrawing | Drawing positioning |
| a | drawingml/2006/main | DrawingML shapes/images |
| pic | drawingml/2006/picture | Picture elements |
| v | urn:schemas-microsoft-com:vml | Legacy VML shapes |
| w14 | wordprocessingml/2010/wordml | Word 2010 extensions |
