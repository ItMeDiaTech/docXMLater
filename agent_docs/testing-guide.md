# Testing Guide

## Test Organization

```
tests/
  ├── core/           # Document, Parser, Generator tests
  ├── elements/       # Paragraph, Run, Table, etc.
  ├── formatting/     # Style, Numbering tests
  ├── helpers/        # Test utilities (validateOoxml, compareDocx, normalizeXml)
  ├── golden/         # Golden file tests (structural regression)
  ├── utils/          # Unit conversion, validation tests
  ├── xml/            # XMLBuilder, XMLParser tests
  └── zip/            # ZipHandler tests
```

## Writing Tests

### Unit Tests

- Test individual methods in isolation
- Mock dependencies when testing internal logic
- Use descriptive test names: `should [expected behavior] when [condition]`

### Integration Tests

- Test document create → modify → save → reload cycles
- Verify round-trip fidelity for load → save operations
- Use `validateOoxml()` to check OOXML compliance of generated documents

### Golden File Tests

- Located in `tests/golden/`
- Compare generated DOCX output against known-good baseline files
- Regenerate baselines: `UPDATE_GOLDEN=true npm test -- --testPathPattern=golden`
- Uses `compareDocx()` to unzip and diff XML parts with normalization

## OOXML Validation

```typescript
import { validateOoxml } from '../helpers/validateOoxml';

it('should produce valid OOXML', async () => {
  const doc = Document.create();
  // ... build document ...
  const buffer = await doc.save();
  await validateOoxml(buffer);
  doc.dispose();
});
```

## Test Naming Conventions

- Files: `[Feature].test.ts` (e.g., `Paragraph.test.ts`)
- Golden tests: `[feature].golden.test.ts`
- Describe blocks match class/function names
- Test names use `should` + expected behavior

## Common Patterns

### Always dispose documents

```typescript
const doc = Document.create();
try {
  // ... test logic ...
} finally {
  doc.dispose();
}
```

### Testing XML output

```typescript
const xml = element.toXML();
expect(xml).toContain('<w:b/>');
expect(xml).not.toContain('<w:rsid');
```

### Testing round-trip

```typescript
const doc = Document.create();
// ... add content ...
const buffer = await doc.save();
const reloaded = await Document.load(buffer);
// ... verify content matches ...
reloaded.dispose();
doc.dispose();
```
