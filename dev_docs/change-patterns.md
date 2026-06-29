# Change Patterns

Common multi-file change patterns. If you modify X, also change Y.

## Adding a New Element Type

1. **Create element class** in `src/elements/NewElement.ts`
   - Extend base or implement required interfaces
   - Add `toXML()` method for serialization
2. **Add to Parser** in `src/core/Parser.ts`
   - Handle the element's XML tag in the parsing switch/case
3. **Add to Generator** in `src/core/Generator.ts`
   - Include element in XML generation output
4. **Export from index** in `src/index.ts`
   - Add to public API exports
5. **Add type guard** if needed (e.g., `isNewElement()`)
6. **Write tests** in `tests/elements/NewElement.test.ts`
7. **Update content types** if element uses new part types

## Adding a New Manager

1. Create `src/managers/NewManager.ts` or `src/formatting/NewManager.ts`
2. Initialize in `Document.create()` and `Parser.parse()`
3. Wire into `Generator.generate()` for save pipeline
4. Add dirty flag tracking if the manager supports modification
5. Export from `src/index.ts`

## Modifying Serialization (toXML)

1. Update the element's `toXML()` method
2. Check corresponding parser handles the new XML
3. Verify round-trip: load → save → load produces same result
4. Run golden file tests: `UPDATE_GOLDEN=true npm test -- golden`
5. Check OOXML validation: tests using `validateOoxml()` helper

## Adding a New Relationship Type

1. Add the relationship type constant to `src/zip/types.ts`
2. Update `[Content_Types].xml` generation in Generator
3. Update relationship file generation in Generator
4. Handle in Parser when loading documents
5. Wire up the rId assignment in the relevant manager

## Modifying Styles/Numbering

1. Update `StylesManager` or `NumberingManager`
2. Set `isModified()` flag so save pipeline regenerates XML
3. Ensure merge methods in Generator handle the changes
4. Verify `_original*Xml` preservation still works (round-trip test)

## Adding/Changing Document Properties

1. Modify `DocumentProperties` type in `src/core/Document.ts`
2. Update `core.xml` generation in Generator
3. Update parsing in Parser
4. Set dirty flag for regeneration
