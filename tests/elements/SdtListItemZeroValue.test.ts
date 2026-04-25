/**
 * `<w:sdtPr><w:dropDownList><w:listItem w:value="0"/>` — SDT list
 * items whose value happens to be a purely-numeric string must
 * round-trip.
 *
 * Per ECMA-376 Part 1 §17.5.2.13 CT_SdtListItem:
 *   - `w:displayText` (ST_String, optional) — user-visible label
 *   - `w:value` (ST_String, required) — internal key
 *
 * Both are xsd:string, so any string is legal. Two compounding bugs
 * dropped legitimate list items silently:
 *   1. XMLParser's `parseAttributeValue: true` coerces purely-numeric
 *      attribute values like `"0"` / `"123"` into JS numbers.
 *   2. The parser's `if (item['@_w:displayText'] && item['@_w:value'])`
 *      truthy gate dropped the item when EITHER attribute was the
 *      coerced number `0` OR the empty string — so a common pattern
 *      like a "None" / "Clear" option with `w:value=""` or a numeric
 *      catalog where `w:value="0"` was the first entry silently
 *      vanished on every load → save round-trip.
 * Storage-side, the emitted XML then broke `ListItem.value: string`
 * because the coerced numbers flowed through the model as raw
 * numbers rather than strings.
 *
 * Iteration 111 gates on presence (`!== undefined`) and coerces both
 * attrs via `String(…)` so numeric and empty values survive intact.
 * It also applies the idiomatic Word fallback of defaulting
 * `displayText` to `value` when displayText is omitted.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { StructuredDocumentTag } from '../../src/elements/StructuredDocumentTag';

async function loadAndReadDropDownItems(
  listItemsXml: string
): Promise<Array<{ displayText: string; value: string }>> {
  const zipHandler = new ZipHandler();
  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:sdt>
      <w:sdtPr>
        <w:id w:val="100"/>
        <w:dropDownList>
          ${listItemsXml}
        </w:dropDownList>
      </w:sdtPr>
      <w:sdtContent>
        <w:p><w:r><w:t>x</w:t></w:r></w:p>
      </w:sdtContent>
    </w:sdt>
    <w:p><w:r><w:t>d</w:t></w:r></w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const sdt = doc
    .getBodyElements()
    .find((el): el is StructuredDocumentTag => el instanceof StructuredDocumentTag);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const items = ((sdt as any)?.properties?.dropDownList?.items ?? []) as Array<{
    displayText: string;
    value: string;
  }>;
  doc.dispose();
  return items;
}

describe('<w:listItem> numeric / empty value round-trip', () => {
  it('preserves w:value="0" as a string "0" (previously coerced to 0 and dropped)', async () => {
    const xml = `
      <w:listItem w:displayText="Zero" w:value="0"/>
      <w:listItem w:displayText="One" w:value="1"/>`;
    const items = await loadAndReadDropDownItems(xml);
    expect(items.length).toBe(2);
    expect(items[0]).toEqual({ displayText: 'Zero', value: '0' });
    expect(items[1]).toEqual({ displayText: 'One', value: '1' });
  });

  it('preserves w:value="" (empty string, e.g. a blank/none option)', async () => {
    const xml = `
      <w:listItem w:displayText="" w:value=""/>
      <w:listItem w:displayText="Option 1" w:value="opt1"/>`;
    const items = await loadAndReadDropDownItems(xml);
    expect(items.length).toBe(2);
    expect(items[0]).toEqual({ displayText: '', value: '' });
    expect(items[1]).toEqual({ displayText: 'Option 1', value: 'opt1' });
  });

  it('defaults displayText to value when w:displayText is omitted', async () => {
    const xml = `<w:listItem w:value="onlyvalue"/>`;
    const items = await loadAndReadDropDownItems(xml);
    expect(items.length).toBe(1);
    expect(items[0]).toEqual({ displayText: 'onlyvalue', value: 'onlyvalue' });
  });

  it('skips items missing the required w:value (spec conformance)', async () => {
    const xml = `
      <w:listItem w:displayText="Missing value"/>
      <w:listItem w:displayText="Valid" w:value="v"/>`;
    const items = await loadAndReadDropDownItems(xml);
    expect(items.length).toBe(1);
    expect(items[0]).toEqual({ displayText: 'Valid', value: 'v' });
  });
});
