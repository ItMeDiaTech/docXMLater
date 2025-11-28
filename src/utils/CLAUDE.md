# Utils Module Documentation

The `utils` module provides utility functions, validation, revision processing, logging, and various helper functions used throughout the framework.

## Module Overview

This module contains standalone utilities that support the core document operations. Key areas include:
- Unit conversions (twips, EMUs, points)
- Revision acceptance and processing
- Document validation and corruption detection
- Logging infrastructure
- Error handling

**Location:** `src/utils/`

**Total Files:** 14 TypeScript files

## File Reference

### Unit Conversions (`units.ts`)

Converts between different measurement units used in DOCX.

**Key Functions:**
```typescript
// Twips (1/20 of a point)
twipsToPoints(twips: number): number
pointsToTwips(points: number): number
twipsToInches(twips: number): number
inchesToTwips(inches: number): number
twipsToCentimeters(twips: number): number
centimetersToTwips(cm: number): number

// EMUs (English Metric Units - used for images)
emusToPoints(emus: number): number
pointsToEmus(points: number): number
emusToInches(emus: number): number
inchesToEmus(inches: number): number
emusToCentimeters(emus: number): number
centimetersToEmus(cm: number): number

// Points
pointsToInches(points: number): number
inchesToPoints(inches: number): number
pointsToPixels(points: number, dpi?: number): number
pixelsToPoints(pixels: number, dpi?: number): number

// Half-points (used for font sizes)
halfPointsToPoints(halfPoints: number): number
pointsToHalfPoints(points: number): number
```

**Common Conversions:**
- 1 inch = 1440 twips
- 1 inch = 914400 EMUs
- 1 inch = 72 points
- 1 point = 20 twips

### Validation (`validation.ts`)

Document and content validation utilities.

**Key Functions:**
```typescript
// Text validation
validateRunText(text: string): { valid: boolean; warnings: string[] }
cleanXmlFromText(text: string): string

// Document validation
isValidDocxStructure(files: string[]): boolean
```

**Use Cases:**
- Detect XML patterns accidentally passed as text content
- Validate DOCX file structure
- Clean user input before adding to document

### Corruption Detection (`corruptionDetection.ts`)

Detects and reports document corruption issues.

**Key Functions:**
```typescript
detectCorruptionInDocument(doc: Document): CorruptionReport
```

**CorruptionReport:**
```typescript
interface CorruptionReport {
  isCorrupted: boolean;
  summary: string;
  locations: CorruptionLocation[];
}
```

**Detection Capabilities:**
- XML markup in text content
- Invalid character sequences
- Malformed element structures
- Suggested fixes for each issue

### Formatting Utilities (`formatting.ts`)

Text and content formatting helpers.

**Key Functions:**
```typescript
normalizeColor(color: string): string  // Uppercase 6-char hex
formatSpacing(spacing: SpacingOptions): object
```

### Deep Clone (`deepClone.ts`)

Safe deep cloning of objects.

```typescript
deepClone<T>(obj: T): T
```

## Revision Processing

### Accept Revisions (`acceptRevisions.ts`)

Main entry point for accepting tracked changes.

**Key Functions:**
```typescript
acceptAllRevisions(
  zipHandler: ZipHandler,
  options?: AcceptRevisionsOptions
): Promise<void>
```

**Options:**
```typescript
interface AcceptRevisionsOptions {
  acceptInsertions?: boolean;  // default: true
  acceptDeletions?: boolean;   // default: true
  acceptMoves?: boolean;       // default: true
  acceptPropertyChanges?: boolean;  // default: true
  useDomParser?: boolean;      // default: true (use RevisionWalker)
}
```

**Processing Steps:**
1. Process document.xml
2. Process all headers and footers
3. Remap image relationship IDs (prevent duplicates)
4. Clean up metadata files (people.xml, settings.xml)
5. Reset revision counts

### RevisionWalker (`RevisionWalker.ts`)

DOM-based tree walker for accepting revisions.

**Advantages over Regex:**
- Handles nested revisions correctly
- Preserves element order via `_orderedChildren`
- More robust against complex document structures
- Prevents ReDoS attacks

**Key Methods:**
```typescript
RevisionWalker.processTree(
  parsed: ParsedObject,
  options?: RevisionWalkerOptions
): ParsedObject
```

**How It Works:**
1. Walks the parsed XML tree depth-first
2. For `w:ins` elements: unwraps content, removes wrapper
3. For `w:del` elements: removes entire element with content
4. For move elements: handles `moveFrom`/`moveTo` pairs
5. For property changes: removes the change tracking elements

### ChangelogGenerator (`ChangelogGenerator.ts`)

Converts revisions to structured changelog data.

**Key Methods:**
```typescript
// Generate entries from document
ChangelogGenerator.fromDocument(
  doc: Document,
  options?: ChangelogOptions
): ChangelogEntry[]

// Output formats
ChangelogGenerator.toJSON(entries: ChangelogEntry[]): string
ChangelogGenerator.toMarkdown(entries: ChangelogEntry[]): string
ChangelogGenerator.toHTML(entries: ChangelogEntry[]): string
ChangelogGenerator.toCSV(entries: ChangelogEntry[]): string

// Analysis
ChangelogGenerator.getTimeline(entries: ChangelogEntry[]): Map<string, ChangelogEntry[]>
ChangelogGenerator.getSummary(entries: ChangelogEntry[]): ChangelogSummary
ChangelogGenerator.consolidate(entries: ChangelogEntry[]): ChangelogEntry[]
```

**Options:**
```typescript
interface ChangelogOptions {
  includeFormattingChanges?: boolean;
  consolidate?: boolean;
  maxContextLength?: number;
  filterAuthors?: string[];
  filterCategories?: ChangeCategory[];
  sortBy?: 'date' | 'author' | 'type';
  sortOrder?: 'asc' | 'desc';
}
```

**Output Formats:**
- **JSON**: Programmatic consumption
- **Markdown**: Human-readable with categories
- **HTML**: Styled with colors and badges
- **CSV**: Spreadsheet-compatible

### SelectiveRevisionAcceptor (`SelectiveRevisionAcceptor.ts`)

Accept or reject revisions by specific criteria.

**Key Methods:**
```typescript
SelectiveRevisionAcceptor.accept(
  doc: Document,
  criteria: SelectionCriteria
): AcceptResult

SelectiveRevisionAcceptor.reject(
  doc: Document,
  criteria: SelectionCriteria
): RejectResult
```

**Selection Criteria:**
```typescript
interface SelectionCriteria {
  authors?: string[];
  types?: RevisionType[];
  dateRange?: { start: Date; end: Date };
  categories?: RevisionCategory[];
}
```

### RevisionAwareProcessor (`RevisionAwareProcessor.ts`)

Handles revisions during document processing operations.

**Processing Modes:**
- `accept_all` - Accept revisions before processing (default)
- `preserve` - Keep revisions, skip conflicting operations
- `preserve_and_wrap` - Keep revisions, wrap conflicts in new tracked changes

### Strip Tracked Changes (`stripTrackedChanges.ts`)

Completely remove all revision markup.

```typescript
stripTrackedChanges(zipHandler: ZipHandler): Promise<void>
```

## Logging Infrastructure

### Logger (`logger.ts`)

Configurable logging for the framework.

**Log Levels:**
```typescript
enum LogLevel {
  DEBUG = 'debug',
  INFO = 'info',
  WARN = 'warn',
  ERROR = 'error'
}
```

**Logger Implementations:**
```typescript
// Console output with timestamps
const consoleLogger = new ConsoleLogger(LogLevel.INFO);

// Silent (no output)
const silentLogger = new SilentLogger();

// Collect logs in memory
const collectingLogger = new CollectingLogger();
```

**Environment Configuration:**
```bash
# Enable debug logging
DEBUG=docxmlater npm test

# Set specific level
DOCXMLATER_LOG_LEVEL=info npm test
DOCXMLATER_LOG_LEVEL=debug npm test
DOCXMLATER_LOG_LEVEL=silent npm test
```

**Programmatic Configuration:**
```typescript
import { setGlobalLogger, ConsoleLogger, LogLevel } from 'docxmlater';

// Enable info-level logging
setGlobalLogger(new ConsoleLogger(LogLevel.INFO));

// Custom logger
setGlobalLogger({
  debug: (msg, ctx) => myLogger.debug(msg, ctx),
  info: (msg, ctx) => myLogger.info(msg, ctx),
  warn: (msg, ctx) => myLogger.warn(msg, ctx),
  error: (msg, ctx) => myLogger.error(msg, ctx),
});
```

**Scoped Logging:**
```typescript
import { createScopedLogger, getGlobalLogger } from 'docxmlater';

const logger = createScopedLogger(getGlobalLogger(), 'MyComponent');
logger.info('Operation completed', { count: 10 });
// Output: 12:34:56.789 [INFO ] [MyComponent] Operation completed count=10
```

**What Gets Logged:**

| Component | INFO Level | DEBUG Level |
|-----------|------------|-------------|
| Document | Load/save operations | + detailed operations |
| ZipHandler | File operations, validation | + file counts, sizes |
| DocumentParser | Parse completion, warnings | + element counts |
| DocumentGenerator | Generation steps | + XML sizes |
| XMLParser | - | Parse operations |
| RevisionManager | Clear, summary stats | + each registration |
| ChangelogGenerator | Generation, consolidation | + filtering details |

### Diagnostics (`diagnostics.ts`)

Development and troubleshooting utilities.

**Key Functions:**
```typescript
// Debug XML structure
debugXmlStructure(xml: string): void

// Element inspection
inspectElement(element: any): string
```

### Error Handling (`errorHandling.ts`)

Custom error types and handling utilities.

**Error Types:**
```typescript
class DocXMLaterError extends Error { }
class InvalidDocxError extends DocXMLaterError { }
class ParseError extends DocXMLaterError { }
class ValidationError extends DocXMLaterError { }
```

## Testing

**Test Files:**
- `tests/utils/validation.test.ts` - Validation functions
- `tests/utils/corruptionDetection.test.ts` - Corruption detection (15+ tests)

**Total: 30+ util-related tests**

## Usage Examples

### Accept All Tracked Changes

```typescript
import { Document } from 'docxmlater';

// Load with automatic acceptance (default)
const doc = await Document.load('input.docx', {
  revisionHandling: 'accept'
});
```

### Generate Changelog

```typescript
import { Document, ChangelogGenerator } from 'docxmlater';

const doc = await Document.load('input.docx', {
  revisionHandling: 'preserve'
});

const entries = ChangelogGenerator.fromDocument(doc, {
  includeFormattingChanges: true,
  sortBy: 'date',
  sortOrder: 'desc'
});

const markdown = ChangelogGenerator.toMarkdown(entries);
console.log(markdown);
```

### Selective Revision Acceptance

```typescript
import { Document, SelectiveRevisionAcceptor } from 'docxmlater';

const doc = await Document.load('input.docx', {
  revisionHandling: 'preserve'
});

// Accept only Alice's insertions from January
const result = SelectiveRevisionAcceptor.accept(doc, {
  authors: ['Alice'],
  types: ['insert'],
  dateRange: {
    start: new Date('2025-01-01'),
    end: new Date('2025-01-31')
  }
});

console.log(`Accepted ${result.accepted} revisions`);
```

### Enable Debug Logging

```typescript
import { setGlobalLogger, ConsoleLogger, LogLevel } from 'docxmlater';

// Enable for troubleshooting
setGlobalLogger(new ConsoleLogger(LogLevel.DEBUG));

// Your document operations will now log details
const doc = await Document.load('input.docx');
```

## See Also

- `src/core/CLAUDE.md` - Document class
- `src/elements/CLAUDE.md` - Element classes including Revision
- `docs/guides/using-track-changes.md` - Track changes guide
- Main `CLAUDE.md` - Debug Logging section
