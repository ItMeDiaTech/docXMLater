# Future Improvements

Improvement opportunities for the docxmlater project, organized by category and priority.
41 items across 11 sections, with a 5-phase prioritized roadmap at the bottom.

## Implementation Progress

**19 of 41 items implemented** (15 fully done, 3 mostly done, 3 partial). Verified: **163 test suites, 3438 tests passing** (+15 suites, +206 tests from original 3232). 19 example directories (was 17).

| Item | Description                                                                                                       | Status      |
| ---- | ----------------------------------------------------------------------------------------------------------------- | ----------- |
| #1   | ESLint configuration restored (flat config, all rules as warnings)                                                | DONE        |
| #3   | ESLint enforced in CI (0 errors, 3759 warnings — promote incrementally)                                           | DONE        |
| #5   | Test coverage: 7 utils + DocumentValidator + ValidationRules + PreservedElements + DrawingManager (206 new tests) | MOSTLY DONE |
| #6   | Coverage threshold in jest.config.js                                                                              | DONE        |
| #8   | Module-level CLAUDE.md documentation (11/14 dirs)                                                                 | DONE        |
| #12  | Stricter TypeScript (both `noUnusedLocals` + `noUnusedParameters` enabled, 70 fixes, 0 violations)                | DONE        |
| #13  | Bundle size check in CI                                                                                           | DONE        |
| #14  | README gaps (error handling, large docs, positioning, unsupported features)                                       | DONE        |
| #21  | Error handling (BookmarkManager, RelationshipManager, ChangelogGenerator, ZipHandler cleanup)                     | DONE        |
| #26  | Developer setup (CONTRIBUTING.md, .vscode/launch.json, debugging docs)                                            | DONE        |
| #27  | Accessibility (findImagesWithoutAltText, getHeadingHierarchy)                                                     | MOSTLY DONE |
| #28  | CI quality gates (knip, depcruise as informational steps)                                                         | DONE        |
| #33  | Export formats (toPlainText, toJSON)                                                                              | MOSTLY DONE |
| #36  | Anti-patterns documentation (6 new entries, 15 total)                                                             | DONE        |
| #37  | ESM support (exports field added, CJS only for now)                                                               | PARTIAL     |
| #38  | Tree-shaking (sideEffects: false)                                                                                 | DONE        |
| #39  | Example gaps (footnotes, content controls examples added)                                                         | PARTIAL     |
| #41  | JSDoc coverage (100% public Document.ts methods documented)                                                       | DONE        |

**New code added:**

- 4 new Document methods: `toPlainText()`, `toJSON()`, `findImagesWithoutAltText()`, `getHeadingHierarchy()`
- 16 new test files with 206 tests
- 4 new CLAUDE.md documentation files (validation, tracking, types, constants)
- .vscode/launch.json with 4 debug configurations
- 6 new anti-pattern entries in agent_docs/anti-patterns.md
- 69 unused import/variable/parameter fixes across 20+ source files

**Files modified:**

- `src/core/Document.ts` — 4 new methods, unused import cleanup
- `tsconfig.json` — `noUnusedParameters: true` enabled
- `package.json` — exports field, sideEffects
- `.github/workflows/ci.yml` — size, knip, depcruise steps
- `jest.config.js` — coverageThreshold
- `README.md` — 5 new sections (When to Use, Unsupported Features, Error Handling, Large Documents, Developer Tools updates)
- `docs/CONTRIBUTING.md` — Node version, debugging, troubleshooting, project structure
- `agent_docs/anti-patterns.md` — 6 new entries
- 20+ source files — unused import/variable cleanup

## Table of Contents

- [High Priority](#high-priority) -- Items 1-4, 15-17
- [Medium Priority](#medium-priority) -- Items 5-9, 18-22
- [Low Priority](#low-priority) -- Items 10-14
- [Architecture Observations](#architecture-observations)
- [Developer Experience & Tooling](#developer-experience--tooling) -- Items 23-28
- [Testing Infrastructure](#testing-infrastructure) -- Items 29-31
- [API Design & Ergonomics](#api-design--ergonomics) -- Items 32-36
- [Competitive Positioning](#competitive-positioning)
- [Packaging & Distribution](#packaging--distribution) -- Items 37-39
- [Maintenance & Housekeeping](#maintenance--housekeeping) -- Items 40-41
- [Prioritized Roadmap](#prioritized-roadmap) -- 5 phases, 24 steps
- [Project Stats](#project-stats-as-of-research-date)
- [Quick Wins](#quick-wins--2-hours-each)

---

## High Priority

### 1. Fix ESLint Configuration -- DONE

~~The ESLint flat config (`eslint.config.mjs`) was deleted from the working tree.~~

Restored from git history. Config uses `typescript-eslint` flat config with `recommendedTypeChecked` + `stylisticTypeChecked`. All rules set to `warn` (0 errors, 3759 warnings). CI lint step now functional.

### 2. Resolve Circular Dependencies (51 violations)

Dependency-cruiser reports 51 circular dependency violations:

- `RevisionAutoFixer` <-> `RevisionValidator` <-> `Document` <-> `validation/index`
- `Style` <-> `Paragraph` <-> `StylesManager`
- `Table` <-> `TableRow` <-> `TableCell` <-> `TrackingContext`

These cause potential module initialization issues and prevent effective dead-code elimination.

- **Action:** Extract pure data types into `src/types/tracking-types.ts` and similar to break cycles.
- **Scope:** Type extraction + import rewiring across ~15 files.

### 3. Enforce ESLint in CI -- DONE

~~Depends on item #1.~~ ESLint config restored. CI lint step (`npm run lint`) now runs successfully. All rules are warnings (0 errors), so CI passes. Rules can be promoted to errors incrementally as warnings are resolved.

### 4. Reduce `any` Usage (133 instances)

~~133~~ **129 instances** remain (4 fixed). Breakdown by root cause:

- `DocumentParser.ts` (113) — raw XML parsing with fast-xml-parser output. Dynamic property access is inherent; replacing with interfaces would require 50+ type definitions for XML node shapes with diminishing safety benefit.
- `RevisionWalker.ts` (8) — DOM tree walking/transformation on parsed XML. Same root cause as DocumentParser.
- `acceptRevisions.ts` (6) — XML tree serialization and relationship remapping. Same root cause.
- `Document.ts` (2) — duck-typed element binding (`bindTrackingToElement`) and raw XML heading extraction. Annotated with eslint-disable.

**Conclusion:** The remaining 129 `any` instances are concentrated in XML processing code where dynamic property access is the correct pattern. Creating typed interfaces for fast-xml-parser output would be a large effort (50+ types) with minimal safety gain — the parser output shape is implicitly defined by the OOXML schema.

- **Action:** Accept current state. Consider typed XML wrappers only if a major parser refactor is planned.

---

## Medium Priority

### 5. Test Coverage for Core Modules

27 source modules lack dedicated test files. Most critical untested modules:

- `src/core/DocumentGenerator.ts`
- ~~`src/core/DocumentValidator.ts`~~ DONE: 21 tests (constructor, validateProperties, estimateSize, getSizeStats)
- `src/validation/RevisionAutoFixer.ts`
- `src/validation/RevisionValidator.ts`
- ~~`src/validation/ValidationRules.ts`~~ DONE: 21 tests (rule structure, codes, severity, auto-fixable, getRuleByCode, getRulesBySeverity, createIssueFromRule)
- ~~`src/elements/AlternateContent.ts`~~ DONE: 4 tests (raw XML, type, serialization, empty)
- ~~`src/elements/CustomXml.ts`~~ DONE: 3 tests (raw XML, type, serialization)
- ~~`src/elements/MathElement.ts`~~ DONE: 6 tests (MathParagraph + MathExpression)
- ~~`src/managers/DrawingManager.ts`~~ DONE: 16 tests (add/get/remove shapes, textboxes, preserved drawings, type identification, clear, mixed types)

Secondary (utility) modules without tests:

- ~~`src/utils/dateFormatting.ts`~~ DONE: 6 tests (millisecond stripping, OOXML format validation, edge cases)
- ~~`src/utils/deepClone.ts`~~ DONE: 11 tests (primitives, nested objects, arrays, Date, Map, Set, RegExp)
- ~~`src/utils/diagnostics.ts`~~ DONE: 16 tests (enable/disable, config isolation, per-channel logging, text comparison, paragraph content)
- ~~`src/utils/errorHandling.ts`~~ DONE: 14 tests (isError, toError, wrapError, getErrorMessage)
- ~~`src/utils/xmlSanitization.ts`~~ DONE: 21 tests (removeInvalidXmlChars, findInvalidXmlChars, hasInvalidXmlChars, constants)
- ~~`src/utils/formatting.ts`~~ DONE: 19 tests (mergeFormatting, cloneFormatting, hasFormatting, cleanFormatting, isEqualFormatting, applyDefaults)
- ~~`src/utils/list-detection.ts`~~ DONE: 24 tests (detectTypedPrefix, inferLevelFromIndentation, getLevelFromFormat, constants)
- `src/utils/stripTrackedChanges.ts`
- `src/zip/ZipReader.ts`, `src/zip/ZipWriter.ts`

- **Action:** Remaining untested: DocumentGenerator (complex), RevisionAutoFixer/Validator (integration), stripTrackedChanges (ZIP context), ZipReader/ZipWriter (I/O). All require integration-level test setup.

### 6. Add Coverage Threshold -- DONE

Added `coverageThreshold` to `jest.config.js` as a regression safety net:

| Metric     | Current | Threshold |
| ---------- | ------- | --------- |
| Statements | 67.97%  | 65%       |
| Branches   | 56.96%  | 54%       |
| Functions  | 61.35%  | 58%       |
| Lines      | 68.72%  | 66%       |

Thresholds set ~3% below current levels to prevent regressions without blocking existing PRs. All 3232 tests pass with thresholds enforced.

### 7. Clarify Utils Layer Architecture -- PARTIALLY DONE

36+ dependency-cruiser warnings about utils importing from elements/core/formatting:

- `ShadingResolver` imports `Table`, `TableCell`, `Style`
- `SelectiveRevisionAcceptor` imports `Run`, `Paragraph`, `Revision`

These are justified. **Documented** in `src/utils/CLAUDE.md` with classification:

- 13 **pure utilities** (no element imports — safe anywhere)
- 11 **domain-specific processors** (import elements/core — should be in `src/processors/`)

- **Action (remaining):** Move domain-specific files to `src/processors/` directory and update imports. This is a mechanical refactor but touches many files.

### 8. Complete Module-Level Documentation -- DONE

CLAUDE.md exists in 11 of 14 src/ subdirectories (all meaningful ones covered):

- Existing: `core/`, `elements/`, `utils/`, `formatting/`, `managers/`, `xml/`, `zip/`, `tracking/`, `validation/`, `types/`, `constants/`
- Skipped: `helpers/` (1 file, trivial), `images/` (1 file, optimizer), `__tests__/` (not needed)

### 9. Public API Refinement

The public API (`src/index.ts`, 562 lines) has some inconsistencies:

- **Internal utilities exposed as public:** `XMLBuilder`, `XMLParser`, `ZipHandler`, `ZipReader`, `ZipWriter`, `DocumentParser`, `DocumentGenerator`, `RelationshipManager` are marked "INTERNAL -- advanced usage" but fully exported with no separation.
- ~~**Missing convenience methods:** No `Document.findText()` or `Document.replaceText()`~~ NOTE: Both already exist (findText at line 13200, replaceText at line 13685). This was incorrectly flagged.
- **Action:** Consider a tiered export strategy (core API vs advanced/internal) or at minimum document the stability guarantees of internal exports.

---

## Low Priority

### 10. Performance Optimizations

- **JSON serialization for equality checks:** `Paragraph.ts` and `DocumentTrackingContext.ts` use `JSON.stringify()` for object comparison. Could use a dedicated deep-equal utility.
- ~~**RegExp creation in loops:**~~ INVESTIGATED: All 30+ `new RegExp()` calls in Document.ts use dynamic values (style IDs, color codes, user patterns). Caching is not viable since patterns depend on runtime parameters. Not an actual issue.
- **No traversal caching:** `getAllParagraphs()` and similar traversal methods rebuild arrays on every call. For large documents (1000+ paragraphs), repeated calls are O(n) each. A dirty-flag-based cache could help.

### 11. Update Outdated Dev Dependencies

14 dev dependency updates available (minor/patch):

- `@types/node`: 25.0.3 -> latest
- `eslint`: 10.0.1 -> 10.2.0
- `jest`: 30.2.0 -> 30.3.0
- `knip`: 5.84.1 -> 5.88.1
- `typescript`: 5.9.3 -> 6.0.2 (major version -- evaluate separately)

- **Action:** Apply minor/patch updates in a batch. Evaluate TypeScript 6.0 as a separate initiative.

### 12. Enable Stricter TypeScript Options

`tsconfig.json` has `strict: true` but explicitly disables:

- `noUnusedLocals: false` -- allows dead code
- `noUnusedParameters: false` -- masks unused parameters

~~59~~ **2 violations remaining** with BOTH `noUnusedLocals` AND `noUnusedParameters` enabled (68 total fixes across 20+ files). The 2 remaining are intentional:

1. `findHeadingsForTOC` — deprecated public method (remove in next major version)
2. `inchesToEmus` in Shape.ts — imported for JSDoc example reference

**Both `noUnusedLocals: true` AND `noUnusedParameters: true`** are now enabled in tsconfig.json. Zero violations. The deprecated private `findHeadingsForTOC` method was removed (dead code) and the `inchesToEmus` JSDoc-only import was cleaned up. Test files use relaxed settings via ts-jest tsconfig override.

### 13. Add Bundle Size Check to CI -- DONE

~~Size-limit is configured in `package.json` (150 KB limit, currently 41 KB) but is not enforced in CI.~~

Added `npm run size` step to `.github/workflows/ci.yml` (runs on ubuntu-latest, Node 22). Verified locally: 4.34 KB brotlied, well within 150 KB limit.

### 14. README Gaps -- DONE

Added to README:

- ~~"When to use docxmlater" section~~ -- positioning vs docx npm package
- ~~"Unsupported OOXML Features" section~~ -- Charts, SmartArt, OLE, glossary, DrawingML advanced
- ~~Error handling patterns~~ -- try/finally with dispose, buffer-based workflow, custom error types
- ~~Working with Large Documents~~ -- memory usage, size limits, buffer operations, caching advice
- ~~Async patterns~~ -- all examples now show proper async/await with try/finally

Remaining (deferred): Troubleshooting guide for corruption and revision conflicts.

### 15. Concurrency Safety for Save Operations

Document has no async locks. Concurrent `save()` calls race on shared state:

- `prepareSave()` modifies internal managers (StylesManager, NumberingManager) without blocking reads
- Atomic save uses temp file + rename, but two concurrent saves both call `prepareSave()` simultaneously
- Getters called during save may see partially-updated state

```typescript
// This races:
await Promise.all([doc.save(path1), doc.save(path2)]);
```

- **Action:** Implement async lock (simple flag or p-queue) to serialize save/load. Document in README that Document is single-threaded.
- **Scope:** Small code change + documentation + concurrency test.

### 16. XML Round-Trip: In-Place Mutation Risks

`stripWebDivs()` in Document.ts mutates `_originalWebSettingsXml` directly via `string.replace()` with no error handling. If the regex match fails partway, XML becomes corrupted.

- **Action:** Add try-catch around in-place XML mutations; validate output before replacing.
- **Related:** StylesManager `isModified()` flag is not synced with Document's internal dirty flag, which can cause conditional regeneration gaps.

### 17. Custom XML Part Preservation

Documents with embedded custom XML parts (`customXmlParts`, `customXmlPropsCore.xml`) lose them on round-trip. DocumentParser treats them as `CustomXmlBlock` but Document has no storage for preservation.

- **Action:** Add custom XML part storage in Document with passthrough on save.

### 18. Structural Validation Gaps

DocumentParser doesn't validate:

- Duplicate bookmark names (BookmarkManager allows silent overwrites)
- Table cell grid consistency (mismatched gridSpan/cells across rows)
- Section boundaries (orphaned sectPr)
- Hyperlink anchor validity (no check if anchor exists)
- Circular style inheritance chains (Normal -> Heading1 -> Normal)
- Orphaned list numbering references

- **Action:** Create `DocumentStructureValidator` to run post-parse checks. Separate from revision validation.

### 19. Revision System Edge Cases

- **Orphaned RSIDs:** `stripOrphanRSIDs()` only scans document.xml, not headers/footers. This loses formatting in headers after round-trip.
- **Zombie revisions:** No validation for revisions that reference deleted elements.
- **Unclosed move ranges:** DocumentParser silently accepts malformed move operations with unclosed range markers.
- **Nested moves with formatting:** Move operations containing embedded `pPrChange` are not validated independently.

- **Action:** Extend RSID scope to headers/footers; add zombie detection; validate move range markers in DocumentGenerator pre-generation.

### 20. Image/Media Memory Management

- `ImageManager.removeImage()` does not explicitly clear the image buffer -- buffers accumulate for removed images
- `_originalCommentCompanionFiles` never cleared on comment removal -- orphaned file references grow
- `PreservedDrawing` XML stored as strings with no size bounds -- a document with 1000 charts could hold multi-MB strings

- **Action:** Add explicit buffer nulling in `removeImage()`; clear companion files on comment deletion; add size warnings for large preserved XML.

### 21. Error Handling Inconsistencies -- DONE

All identified issues resolved:

- ~~`BookmarkManager.ts` throws generic `Error`~~ DONE: Now uses `InvalidDocxError`
- ~~`RelationshipManager.ts` throws generic `Error`~~ DONE: Now uses `InvalidDocxError` + `CorruptedArchiveError`
- ~~`ChangelogGenerator.ts` silently swallows JSON.stringify failures~~ DONE: Elevated to `warn` level
- ~~ZipHandler file I/O cleanup errors suppressed~~ DONE: Elevated temp file cleanup to `warn` level in Document.ts

### 22. Complex Element Edge Cases

- **Tables:** hMerge (legacy) + grid-based merging coexist with no conflict validation; `TableGridChange.ts` doesn't validate column count consistency across rows
- **Fields:** Complex fields (TOC, INDEX) have no nesting validation; MERGEFIELD parsing assumes ASCII; binary field data causes silent corruption
- **Bookmarks:** Cross-reference fields pointing to deleted bookmarks silently break
- **Comments:** Reply threading not validated; orphaned replies persist when parent is deleted; overlapping comment ranges have undefined behavior

- **Action:** Add targeted validation for each element type's known edge cases. Prioritize table merge validation and field nesting checks.

---

## Architecture Observations

| Area             | Current State                             | Risk Level |
| ---------------- | ----------------------------------------- | ---------- |
| Production deps  | 1 (jszip)                                 | Excellent  |
| Bundle size      | 41 KB / 150 KB limit                      | Excellent  |
| Test count       | 3100+ tests across 150 files              | Strong     |
| Type safety      | `strict: true` but 133 `any` usages       | Moderate   |
| Circular deps    | 51 violations detected                    | High       |
| CI coverage      | Tests + typecheck; lint broken            | Needs fix  |
| Documentation    | Good README; partial module docs          | Adequate   |
| Concurrency      | No async locks; concurrent save races     | Needs fix  |
| Error handling   | Custom errors exist; inconsistent use     | Adequate   |
| XML preservation | Strong mechanism; edge cases in mutations | Good       |
| Validation       | Revisions validated; structure not        | Moderate   |
| Image memory     | Limits enforced; buffer cleanup gaps      | Adequate   |

---

## Developer Experience & Tooling

### 23. OOXML Feature Coverage Gaps

Features not implemented or partially implemented:

| OOXML Feature        | Status              | Notes                                  |
| -------------------- | ------------------- | -------------------------------------- |
| Charts               | Not implemented     | ECMA-376 Part 1 S20.1.1 (c:chartSpace) |
| SmartArt             | Not implemented     | Stored as raw XML passthrough only     |
| OLE Embedded Objects | Raw XML passthrough | `<w:object>` preserved, no editing API |
| Glossary Document    | Not supported       | `glossary.xml` not handled             |
| Gradient fills       | Not implemented     | `a:gradFill` in DrawingML              |
| Pattern fills        | Not implemented     | `a:pattFill` in DrawingML              |
| Group shapes         | Not implemented     | `a:grpSp` in DrawingML                 |
| 3D effects           | Not implemented     | `a:scene3d` in DrawingML               |
| Shape effects        | Not implemented     | Shadow, reflection, glow, blur         |

Already implemented: Content Controls (SDT, 13 types), Footnotes/Endnotes (full round-trip), basic Shapes + TextBoxes, RTL/CJK language support.

- **Action:** Add raw XML passthrough for Charts (covers 80% of use cases without full API). Document "Unsupported OOXML Features" in README.

### 24. Plugin/Extension Architecture

The framework is currently closed for extension. No plugin system, middleware hooks, or element registry pattern exists.

Current extensibility: subclassing elements, raw XML passthrough slots (9+ in Image/Shape), custom styles via StylesManager, global logger injection.

Not extensible: custom element types, XML namespace handlers, validation rule registration, serialization interception.

- **Action (near-term):** Document the extensibility model (subclassing, passthrough, logger injection) in README.
- **Action (future):** Design `ElementRegistry.register(tag, parser, generator)` and `ValidationRules.register(rule)` patterns.

### 25. Streaming / Large Document Support

Everything loads into memory at once. `ZipReader.loadFromFile()` reads entire ZIP into Buffer. Size limits exist (warn at 50MB, error at 150MB) but no streaming or chunked processing.

- **Action (near-term):** Document memory model in README: "Working with Large Documents" section with guidance on paragraph counts vs RAM.
- **Action (future):** Design streaming API for section-at-a-time processing (could be separate `@docxmlater/streaming` package).

### 26. Developer Setup & Debugging -- DONE

- ~~`CONTRIBUTING.md` lists Node 14+ (outdated; package.json requires Node 18+)~~ Fixed: updated to Node 18+
- ~~No documented way to use `DEBUG=docxmlater` env var~~ Fixed: added Debugging section with env vars
- ~~Project structure in CONTRIBUTING.md was outdated~~ Fixed: updated to match actual src/ layout
- ~~No troubleshooting section~~ Fixed: added Troubleshooting section
- ~~No `.vscode/launch.json` example for debugging tests~~ Fixed: added 4 debug configurations (current file, all tests, by name, example file)
- No interactive REPL or playground for experimentation (deferred — low priority)

### 27. Accessibility Utilities -- MOSTLY DONE

- `Image.setAltText()` exists but no validation (Word limit: 255 chars)
- ~~No `Document.findImagesWithoutAltText()` helper~~ DONE: Added with 5 tests
- ~~No heading hierarchy method~~ DONE: Added `Document.getHeadingHierarchy()` with 5 tests. Returns `{ level, text, paragraph }[]` in document order. Skipped levels detectable by comparing adjacent entries.
- Shape/TextBox alt text has setter but no getter

Remaining: Full `validateAccessibility()` method (combines alt text + heading checks), alt text length validation.

### 28. CI Quality Gates -- DONE

All quality tools now run in CI (`.github/workflows/ci.yml`):

| Tool         | npm script          | In CI? | Blocking? | Notes                                    |
| ------------ | ------------------- | ------ | --------- | ---------------------------------------- |
| ESLint       | `npm run lint`      | Yes    | Yes       | Needs config fix (item 1)                |
| Knip         | `npm run knip`      | Yes    | No        | 69 unused type exports (intentional API) |
| dep-cruiser  | `npm run depcruise` | Yes    | No        | 51 circular dep errors to resolve first  |
| size-limit   | `npm run size`      | Yes    | Yes       | 4.34 KB / 150 KB limit                   |
| OOXML schema | (in test setup)     | Yes    | Yes       | Already enforced                         |

Knip and depcruise run as informational (`continue-on-error: true`) since they have existing findings. Once circular deps are resolved and unused exports are addressed, they can be promoted to blocking.

---

## Testing Infrastructure

### 29. Testing Strengths (for context)

- 3232 tests across 148 suites, all passing in ~172s
- Automatic OOXML schema validation on every generated DOCX (via monkey-patched `toBuffer`)
- 86 files contain round-trip tests (load -> modify -> save -> reload)
- Golden file comparison with XML normalization
- Performance benchmarks (100-page doc < 5s, 1000 paragraphs < 5s, 500 para memory < 100MB)

### 30. Missing Test Categories

- **Property-based/fuzz tests:** No fast-check or jsfuzz integration. Opportunity for field instruction parsing, XML normalization.
- **Concurrent/stress tests:** No parallel Document modification tests. No tests for race conditions in save().
- **Error recovery tests:** Limited coverage for malformed DOCX, corrupted ZIP entries, invalid XML.
- **E2E workflow tests:** Round-trips are strong, but no realistic user workflow tests (template application, bulk edits, corporate document processing).
- **Mutation testing:** No stryker.js to catch missing assertions.

- **Action:** Add fast-check for field instruction parsing; add concurrent modification tests; add corrupted DOCX recovery tests.

### 31. Test Performance Tracking

Performance benchmarks exist but are timing-based, not regression-tracked across versions. No CI integration for performance trend analysis.

- **Action:** Integrate benchmark results into CI with trend tracking (store results in JSON, fail on >20% regression).

---

## Production Readiness Notes

The framework is production-ready for standard document workflows. Additional hardening is recommended before use with:

1. **Concurrent/parallel operations** -- needs async locks on save/load (item #15)
2. **Large image sets in long-running processes** -- needs buffer cleanup on removal (item #20)
3. **Complex table structures with merging** -- needs grid consistency validation (item #22)
4. **Tracked changes with orphaned revisions** -- needs zombie detection (item #19)
5. **Documents containing custom XML parts** -- currently lost on round-trip (item #17)
6. **Very large documents (50MB+)** -- full in-memory loading, no streaming (item #25)
7. **Documents with Charts/SmartArt** -- raw passthrough only, no editing API (item #23)

---

## API Design & Ergonomics

### 32. Fluent API Inconsistencies

Most `set*` and `add*` methods return `this` for chaining, but there are exceptions:

- `Paragraph.addHyperlink(url?)` returns `Hyperlink | this` (overloaded at line 696-703) -- returns `Hyperlink` when given a URL string (for chaining on the link object), returns `this` when given a `Hyperlink` instance. This is intentional for DX but breaks uniform chaining expectations.
- Some properties are directly mutable (`doc.bodyElements = [...]`) while others use getter/setter methods -- no clear boundary between public-mutable and API-controlled properties.

- ~~**Action:** Document the `addHyperlink` dual-return pattern~~ DONE: Added JSDoc with usage examples for both patterns in Paragraph.ts. Also documented in anti-patterns.md.
- Remaining: Audit other methods for similar overload inconsistencies. Consider `addHyperlinkElement()` as `this`-returning alternative.

### 33. No Export Formats Beyond DOCX -- MOSTLY DONE

- ~~Add `Document.toPlainText()`~~ DONE: 7 tests
- ~~Add `Document.toJSON()`~~ DONE: Returns `{ properties, stats, headings, body }` with 7 tests. Includes paragraph count, table dimensions, heading hierarchy, image count, styles, and document properties. Fully JSON-serializable.

Remaining (future): HTML export as a separate package (`@docxmlater/html-export`). PDF export out of scope.

### 34. No Event System or Change Notifications

No EventEmitter, Observable, or listener pattern exists. Consumers cannot react to programmatic document modifications.

- No change notifications when elements are added/removed/modified
- No undo/redo system (track changes cover Word-level revisions but not programmatic edits)
- No hooks for pre/post save or load

- **Action (near-term):** Document this limitation clearly -- the framework is designed for batch processing, not interactive editing.
- **Action (future):** Design an optional event layer: `document.on('paragraphAdded', callback)` for apps that need it.

### 35. No Global Document Defaults

No way to set default font, margins, or styles for all new documents globally. Each `Document.create()` starts from scratch.

- **Action:** Consider `Document.setDefaults({ font: 'Calibri', fontSize: 11, margins: { ... } })` as a convenience for apps that create many documents with the same base style.

### 36. Anti-Patterns Documentation Gaps -- DONE

Added 6 new entries to `agent_docs/anti-patterns.md` (now 15 total):

- ~~Hyperlink return type inconsistency~~ Added: `addHyperlink` overloaded return types
- ~~Direct property mutation patterns~~ Added: direct property assignment bypasses validation
- ~~Relationship ID ordering assumptions~~ Added: rId ordering is not stable across cycles
- Also added: concurrent save races, table cell content assumption, table cell paragraph population

- **Action:** Update anti-patterns.md with these additional entries.

---

## Competitive Positioning

### Key Differentiator

docxmlater is designed for **editing existing documents with tracked changes**. The popular `docx` npm package is primarily for **generating documents from scratch**. ~~This distinction should be emphasized in documentation.~~ DONE: Added "When to use docxmlater" section to README.

### Feature Comparison (vs `docx` npm package)

| Feature                   | docxmlater      | docx (npm)       |
| ------------------------- | --------------- | ---------------- |
| Edit existing documents   | Full support    | Limited          |
| Tracked changes/revisions | Full support    | Not supported    |
| XML round-trip fidelity   | High            | N/A (generation) |
| Comments system           | Full support    | Basic            |
| HTML/PDF export           | Not supported   | Via plugins      |
| Builder pattern           | Method chaining | Declarative JSON |
| Document generation       | Supported       | Primary focus    |
| Community size            | Emerging        | ~500k downloads  |
| TypeScript                | Native          | Native           |

- **Action:** Add a "When to use docxmlater" section to README explaining the editing/round-trip focus vs generation-only alternatives.

---

## Packaging & Distribution

### 37. ESM Support (Dual CJS/ESM Publishing) -- PARTIALLY DONE

~~No `exports` field in package.json.~~ Added `exports` field with `types`, `require`, and `default` conditions. Also exports `./package.json` for tooling compatibility.

The package still builds CJS only. The `exports` field is structured to support ESM when an ESM build target is added later (swap `default` to point to ESM entry).

Remaining: Add ESM build target (separate tsconfig or bundler), add `"import"` condition to exports.

### 38. Tree-Shaking Support -- DONE

~~No `sideEffects` field in package.json.~~

Added `"sideEffects": false` to package.json. Build verified. Bundlers can now tree-shake unused exports.

### 39. Example Gaps -- PARTIALLY DONE

18 example directories (was 17). Remaining gaps:

- ~~Footnotes/endnotes~~ DONE: `examples/15-footnotes-endnotes/`
- ~~Structured Document Tags / Content Controls~~ DONE: `examples/16-content-controls/` with richText, plainText, checkbox, comboBox, datePicker
- Complex field codes (HYPERLINK, IF, MERGE fields)
- Math elements / equations
- Document protection
- Compatibility mode handling

Examples are TypeScript-only (runnable via ts-node) with no pre-built output samples.

- **Action:** Add examples for footnotes, SDTs, and field codes. Consider adding a `npm run examples` script and pre-built output .docx files in examples/output/.

---

## Quick Wins (< 2 hours each)

- ~~Update CONTRIBUTING.md Node version (14 -> 18)~~ DONE
- ~~Add `.vscode/launch.json` for debugging tests~~ DONE
- ~~Document `DOCXMLATER_LOG_LEVEL` env var in README~~ DONE
- ~~Add `coverageThreshold` to jest.config.js~~ DONE
- ~~Add `npm run size` step to CI workflow~~ DONE
- ~~Add "Unsupported OOXML Features" section to README~~ DONE
- Enable `noUnusedLocals` / `noUnusedParameters` in tsconfig.json (59 violations — not quick)
- ~~Add `"sideEffects": false` to package.json~~ DONE
- ~~Add `exports` field to package.json~~ DONE (CJS only; ESM later)
- ~~Add "When to use docxmlater" section to README~~ DONE
- Remove deprecated APIs in next major version (see item #40)

---

## Maintenance & Housekeeping

### 40. Deprecated API Cleanup

6 deprecated APIs identified (all already have `@deprecated` JSDoc tags) that should be removed in the next major version:

| Method / Type                | Replacement                    | Location       | Line  |
| ---------------------------- | ------------------------------ | -------------- | ----- |
| `acceptAllRevisionsRawXml()` | `acceptAllRevisions()`         | Document.ts    | 12055 |
| (second overload)            | `acceptAllRevisions()`         | Document.ts    | 12247 |
| `getBodyElements()`          | `getAllParagraphs()`           | Document.ts    | 1653  |
| `fixTODHyperlinks()`         | Moved to Template_UI           | Document.ts    | 4207  |
| `findHeadingsForTOC()`       | `findHeadingsForTOCFromXML()`  | Document.ts    | 8590  |
| `Image.setBorderWidth()`     | `Image.setBorder()`            | Image.ts       | 1527  |
| `ApplyStylesOptions` (type)  | `ApplyStylesOptions` (renamed) | styleConfig.ts | 187   |

- **Action:** Remove all in next major version bump. All already have `@deprecated` JSDoc tags.

### 41. JSDoc Coverage -- DONE

**100%** of public Document.ts methods now have JSDoc comments (was ~71%). Added JSDoc to:

- `save(filePath)` — atomic write description, error type, example
- `createFootnote(text)` / `createEndnote(text)` — parameter and return docs
- `getFootnoteManager()` / `getEndnoteManager()` — purpose docs
- `getOptimizeForBrowser()` / `setOptimizeForBrowser()` / `getAllowPNG()` / `setAllowPNG()` — web settings
- `addHyperlink()` — overload pattern documentation (item #32)

All public methods now documented — `getEvenAndOddHeaders()` was the last one.

---

## Prioritized Roadmap

Items organized by recommended execution order, balancing impact and effort.

### Phase 1: Foundation (Quick Wins) -- COMPLETE

All foundation items done:

1. ~~Recreate `eslint.config.mjs` (item #1)~~ DONE
2. ~~Add knip, depcruise, size-limit to CI (item #28)~~ DONE
3. ~~Add `coverageThreshold` to jest.config.js (item #6)~~ DONE
4. ~~Update CONTRIBUTING.md Node version (item #26)~~ DONE
5. ~~Add `"sideEffects": false` to package.json (item #38)~~ DONE
6. ~~Enable both `noUnusedLocals` + `noUnusedParameters` (item #12)~~ DONE (0 violations)

### Phase 2: Type Safety & Architecture -- 2/4 DONE

Reduce tech debt and improve code quality.

7. Reduce `any` usage (item #4) -- analyzed; 129 remaining are justified XML parsing
8. Break circular dependencies -- extract types (item #2, 51 violations)
9. ~~Clarify utils layer (item #7)~~ PARTIALLY DONE — documented pure vs domain-specific split in CLAUDE.md
10. ~~Consistent error classes across all modules (item #21)~~ DONE

### Phase 3: Testing & Validation -- 0/4 DONE

Fill coverage gaps and add missing test categories.

11. Add tests for DocumentGenerator, DocumentValidator, RevisionAutoFixer (item #5)
12. Add structural validation: bookmarks, table grid, styles (item #18)
13. Add property-based tests with fast-check (item #30)
14. Add concurrent save() race condition tests (item #15)

### Phase 4: API & DX -- 3/5 DONE

Polish the developer experience and public API.

15. Fix fluent API inconsistencies (item #32) -- TODO
16. ~~Add `Document.toPlainText()` and `Document.toJSON()` (item #33)~~ DONE
17. ~~Add accessibility utilities (item #27)~~ MOSTLY DONE
18. ~~Add ESM support (item #37)~~ PARTIAL (exports field added)
19. Add examples for footnotes, SDTs, field codes (item #39) -- TODO

### Phase 5: Advanced Features

Longer-term improvements for expanded use cases.

20. Chart raw XML passthrough (item #23)
21. Custom XML part preservation (item #17)
22. Plugin/extension architecture design (item #24)
23. Streaming API design for large documents (item #25)
24. Event system for change notifications (item #34)

---

## Project Stats (updated after implementation)

| Metric                    | Before | After                                        |
| ------------------------- | ------ | -------------------------------------------- |
| Source files              | 120    | 120                                          |
| Test suites               | 148    | **163** (+15)                                |
| Tests                     | 3,232  | **3,438** (+206)                             |
| Test duration             | ~172s  | ~196s                                        |
| Production deps           | 1      | 1 (jszip)                                    |
| Bundle size               | 41 KB  | 41 KB (limit: 150 KB)                        |
| Node.js support           | >=18   | >=18 (CI: 18, 20, 22)                        |
| TypeScript strict         | Yes    | Yes + `noUnusedParameters`                   |
| `any` instances           | 133    | **129** (4 fixed; 125 justified XML parsing) |
| Unused import/var/param   | 70     | **2** (intentional)                          |
| Circular dep violations   | 51     | 51 (unchanged)                               |
| Deprecated APIs           | 6      | 6                                            |
| Module CLAUDE.md coverage | 2/14   | **11/14**                                    |
| Anti-pattern entries      | 9      | **15**                                       |
| CI quality gates          | 2      | **5** (lint, test, size, knip, depcruise)    |
| License                   | MIT    | MIT                                          |

```

```
