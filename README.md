# DOCX Header Line Break Processor

A TypeScript utility using the docXMLater framework to automatically insert line breaks after Header 2 elements within 1x1 tables in Microsoft Word documents.

## Understanding Bullet Points in DOCX/XML

### Structure Overview

Bullet points in DOCX files involve two main components:

1. **Numbering Definitions** (`word/numbering.xml`)
   ```xml
   <w:abstractNum w:abstractNumId="1">
     <w:lvl w:ilvl="0">
       <w:numFmt w:val="bullet"/>
       <w:lvlText w:val="•"/>
       <w:lvlJc w:val="left"/>
     </w:lvl>
   </w:abstractNum>
   ```

2. **Paragraph References** (`word/document.xml`)
   ```xml
   <w:p>
     <w:pPr>
       <w:numPr>
         <w:ilvl w:val="0"/>
         <w:numId w:val="1"/>
       </w:numPr>
     </w:pPr>
     <w:r>
       <w:t>Bullet point text</w:t>
     </w:r>
   </w:p>
   ```

### Common Windows Bullet Symbols

| Symbol | Unicode | Name | Usage |
|--------|---------|------|-------|
| • | U+2022 | Bullet | Default bullet |
| ○ | U+25CB | White Circle | Secondary level |
| ▪ | U+25AA | Black Square | Tertiary level |
| ▫ | U+25AB | White Square | Alternative |
| ◆ | U+25C6 | Black Diamond | Emphasis |
| ➢ | U+27A2 | Arrow | Direction/action |
| ✓ | U+2713 | Check Mark | Completed items |

## Features

- ✅ Detects Header 2 (Heading2 style) within 1x1 tables
- ✅ Checks for existing line breaks between table and next element
- ✅ Inserts line break only when needed
- ✅ Preserves document structure and formatting
- ✅ Supports both low-level XML manipulation and high-level API

## Installation

```bash
# Clone or create the project
mkdir docx-processor
cd docx-processor

# Install dependencies
npm install docxml jszip
npm install -D typescript ts-node @types/node

# Or using the provided package.json
npm install
```

## Usage

### Command Line

```bash
# Basic usage
ts-node process-headers-in-tables.ts input.docx

# With custom output file
ts-node process-headers-in-tables.ts input.docx output.docx

# Verbose mode
ts-node process-headers-in-tables.ts input.docx output.docx --verbose
```

### As a Module

```typescript
import { HeaderTableProcessor } from './process-headers-in-tables';

const processor = new HeaderTableProcessor({
    inputFile: 'document.docx',
    outputFile: 'processed.docx',
    verbose: true
});

await processor.process();
```

## How It Works

### Detection Logic

1. **Table Identification**: Finds all `<w:tbl>` elements
2. **1x1 Verification**: Counts rows (`<w:tr>`) and cells (`<w:tc>`)
3. **Header 2 Check**: Looks for `<w:pStyle w:val="Heading2">`
4. **Gap Analysis**: Examines content after table for existing breaks

### Line Break Insertion

The processor inserts an empty paragraph when:
- Table is exactly 1x1
- Contains Header 2 style
- No line break exists after table
- Next element is not a section break

### XML Structure Added

```xml
<!-- Empty paragraph for line break -->
<w:p w:rsidR="00AB12CD" w:rsidRDefault="00AB12CD">
  <w:pPr>
    <w:spacing w:after="0" w:before="0" w:line="240" w:lineRule="auto"/>
  </w:pPr>
</w:p>
```

## Architecture

### Low-Level API (ZipHandler)
- Direct XML manipulation
- Full control over document structure
- Best for complex transformations

### High-Level API (Document)
- Object-oriented approach
- Type-safe operations
- Simpler for basic edits

## Examples

### Example 1: Processing Multiple Files

```typescript
const files = ['doc1.docx', 'doc2.docx', 'doc3.docx'];

for (const file of files) {
    const processor = new HeaderTableProcessor({
        inputFile: file,
        verbose: false
    });
    await processor.process();
    console.log(`Processed: ${file}`);
}
```

### Example 2: Custom Processing Logic

```typescript
class CustomProcessor extends HeaderTableProcessor {
    protected createEmptyParagraph(): string {
        // Custom spacing or formatting
        return `<w:p>
            <w:pPr>
                <w:spacing w:after="200" w:before="100"/>
            </w:pPr>
        </w:p>`;
    }
}
```

## Troubleshooting

### Document Won't Open
- Validate XML syntax
- Check for unclosed tags
- Verify RSID format (8 hex characters)

### Line Breaks Not Appearing
- Confirm Header 2 style name matches exactly
- Check table structure (must be 1x1)
- Verify output file is being saved

### Performance Issues
- Use buffer operations for large files
- Process in batches for multiple documents
- Consider streaming for files > 10MB

## Testing

Create a test document with:
1. Regular paragraphs
2. 1x1 table with Header 2
3. 2x2 table with Header 2 (should be ignored)
4. 1x1 table without Header 2 (should be ignored)

Run the processor and verify only the 1x1 table with Header 2 gets a line break.

## Dependencies

- `docxml` (docXMLater framework) - TypeScript DOCX manipulation
- `jszip` - ZIP file handling
- `typescript` - TypeScript compiler
- `ts-node` - TypeScript execution

## License

MIT

## Contributing

1. Fork the repository
2. Create your feature branch
3. Test your changes
4. Submit a pull request

## Notes

- The docXMLater framework is accessed via npm package `docxml`
- Original repository: https://github.com/wvbe/docxml
- This implementation uses low-level ZIP/XML manipulation for precise control
- RSID generation ensures Word tracks changes properly
