# ZIP Adapter Abstraction Design

## Purpose

This document outlines a future architecture for abstracting the ZIP handling layer in DocXML. Currently, the framework is tightly coupled to JSZip. This design provides a migration path if JSZip becomes unmaintained, has security issues, or performance needs change.

## Current Status (Phase 2)

**Current Implementation:**
- Direct dependency on JSZip v3.10.1
- `ZipHandler`, `ZipReader`, and `ZipWriter` classes directly use JSZip API
- ~1,000 lines of code coupled to JSZip

**Known Limitations:**
- JSZip loads entire archive into memory (no streaming)
- No incremental save support
- Last major update: June 2023
- High memory usage for large files (>100MB)

## Proposed Architecture

### 1. Interface Definition

```typescript
/**
 * Abstract interface for ZIP operations
 * Allows swapping implementations without changing DocXML code
 */
interface ZipAdapter {
  /**
   * Load ZIP archive from file
   */
  loadFromFile(filePath: string, options?: LoadOptions): Promise<void>;

  /**
   * Load ZIP archive from buffer
   */
  loadFromBuffer(buffer: Buffer, options?: LoadOptions): Promise<void>;

  /**
   * Get file from archive
   */
  getFile(path: string): ZipFile | undefined;

  /**
   * Get all files
   */
  getAllFiles(): FileMap;

  /**
   * Add file to archive
   */
  addFile(path: string, content: string | Buffer, options?: AddFileOptions): void;

  /**
   * Remove file from archive
   */
  removeFile(path: string): boolean;

  /**
   * Save archive to file
   */
  saveToFile(filePath: string, options?: SaveOptions): Promise<void>;

  /**
   * Generate archive as buffer
   */
  toBuffer(options?: SaveOptions): Promise<Buffer>;

  /**
   * Clear all files
   */
  clear(): void;

  /**
   * Validate DOCX structure
   */
  validate(): void;
}
```

### 2. Implementation Adapters

#### JSZip Adapter (Current)
```typescript
class JSZipAdapter implements ZipAdapter {
  private jszip: JSZip;

  // Wraps current JSZip implementation
  // Minimal changes to existing code
}
```

#### Yauzl/Yazl Adapter (Streaming Alternative)
```typescript
class YauzlAdapter implements ZipAdapter {
  // Uses yauzl for reading (streaming)
  // Uses yazl for writing (streaming)
  // Better memory efficiency for large files
}
```

#### Node Native Adapter (Future)
```typescript
class NodeZipAdapter implements ZipAdapter {
  // Uses native Node.js ZIP support when available
  // Potentially best performance
}
```

### 3. Adapter Selection

```typescript
/**
 * Factory for creating ZIP adapters
 */
class ZipAdapterFactory {
  /**
   * Create adapter based on configuration or auto-detection
   */
  static create(type?: 'jszip' | 'yauzl' | 'native'): ZipAdapter {
    if (type === 'yauzl') {
      return new YauzlAdapter();
    }
    if (type === 'native' && isNodeZipAvailable()) {
      return new NodeZipAdapter();
    }
    // Default to JSZip for backwards compatibility
    return new JSZipAdapter();
  }
}
```

### 4. Integration with ZipHandler

```typescript
export class ZipHandler {
  private adapter: ZipAdapter;

  constructor(adapterType?: 'jszip' | 'yauzl' | 'native') {
    this.adapter = ZipAdapterFactory.create(adapterType);
  }

  // All methods delegate to adapter
  async load(filePath: string, options?: LoadOptions): Promise<void> {
    return this.adapter.loadFromFile(filePath, options);
  }

  // ... other methods delegate similarly
}
```

## Migration Path

### Phase 1: Create Abstraction (No Breaking Changes)
1. Define `ZipAdapter` interface
2. Create `JSZipAdapter` wrapping current implementation
3. Update `ZipHandler` to use adapter internally
4. **Result:** Same functionality, prepared for future changes

### Phase 2: Add Alternative Implementations
1. Implement `YauzlAdapter` for streaming support
2. Add configuration option to select adapter
3. Benchmark performance differences
4. **Result:** Users can opt-in to better performance

### Phase 3: Optimize Default
1. Based on benchmarks, choose best default
2. Deprecate JSZip if better alternative found
3. Provide migration guide
4. **Result:** Better performance out-of-the-box

## Benefits

### Flexibility
- Easy to switch implementations if JSZip has issues
- Can optimize for different use cases (small vs. large files)
- Not locked into single dependency

### Performance
- Yauzl/Yazl provides streaming for large files
- Native implementation could be fastest
- Can benchmark and choose best option

### Maintenance
- If JSZip becomes unmaintained, easy to migrate
- Security issues in one adapter don't block entire framework
- Can support multiple adapters simultaneously

### Testing
- Can mock adapter for unit tests
- Easier to test ZIP operations in isolation
- Adapter-specific tests separate from DocXML logic

## Risks & Mitigation

### Risk: Breaking Changes
**Mitigation:**
- Phase 1 maintains 100% backwards compatibility
- Old API continues to work
- Migration is opt-in initially

### Risk: Increased Complexity
**Mitigation:**
- Adapter interface is simple (10-15 methods)
- Most code stays the same
- Only ZipHandler changes significantly

### Risk: Performance Regression
**Mitigation:**
- Benchmark before switching defaults
- Allow users to choose adapter
- Keep JSZip adapter available

## Alternative Approaches Considered

### 1. Stay with JSZip Only
**Pros:** No work required, stable
**Cons:** Locked-in, memory issues persist, potential security risk

### 2. Direct Migration to Different Library
**Pros:** Simpler than abstraction
**Cons:** Risky, no fallback, hard to change later

### 3. Build Custom ZIP Handler
**Pros:** Full control, optimal for DOCX
**Cons:** Massive effort, reinventing wheel, maintenance burden

## Recommendation

**Implement Phase 1 in Phase 3 or 4 of DocXML development**

- Low risk (no breaking changes)
- Future-proofs the framework
- Enables performance improvements later
- Small effort (~2-3 days)

**Timeline:**
- Phase 1: After Phase 3 features complete (table support)
- Phase 2: Based on user feedback about large file handling
- Phase 3: Only if JSZip shows problems

## Example Usage After Migration

```typescript
// Default (JSZip)
const doc = Document.create();
await doc.save('output.docx');

// Streaming for large files
const doc = Document.create({ zipAdapter: 'yauzl' });
await doc.save('large-output.docx'); // Better memory usage

// Native (if available)
const doc = Document.create({ zipAdapter: 'native' });
await doc.save('output.docx'); // Potentially fastest
```

## Conclusion

The ZIP adapter abstraction provides insurance against dependency issues while enabling future performance improvements. Implementation in Phase 1 is low-risk and sets up the framework for long-term success.

**Status:** Design documented, implementation deferred to Phase 3+

---

**Document Version:** 1.0
**Last Updated:** October 2025
**Author:** DocXML Team
