/**
 * `<w:hyperlink>` — numeric-looking string attributes must round-trip
 * as strings (type-contract safety).
 *
 * Per ECMA-376 Part 1 §17.16.22 CT_Hyperlink, `w:anchor`, `w:tooltip`,
 * `w:tgtFrame`, `w:docLocation`, and `r:id` are all `ST_String`
 * (xsd:string). XMLParser's `parseAttributeValue: true` coerces
 * purely-numeric strings (e.g., a bookmark anchor like `"12345"`, or
 * a doc-location reference `"42"`) to JS numbers.
 *
 * The previous parser stored the raw coerced value on the Hyperlink
 * instance. `Hyperlink`'s fields are all declared `string` — storing
 * numbers violated the type contract and would break any downstream
 * code that called `.startsWith(...)`, `.toLowerCase()`, etc. on the
 * anchor / tooltip / tgtFrame / docLocation.
 *
 * Iteration 125 casts every hyperlink-attribute read through
 * `String(...)` so the declared TypeScript contract holds.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { Hyperlink } from '../../src/elements/Hyperlink';

async function loadFirstHyperlink(hlAttrs: string): Promise<Hyperlink | undefined> {
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="1" w:name="12345"/>
      <w:r><w:t>target</w:t></w:r>
      <w:bookmarkEnd w:id="1"/>
      <w:hyperlink${hlAttrs}>
        <w:r><w:t>link</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const paragraph = doc.getParagraphs()[0]!;
  const hyperlink = paragraph.getContent().find((el): el is Hyperlink => el instanceof Hyperlink);
  doc.dispose();
  return hyperlink;
}

describe('<w:hyperlink> numeric attribute type-contract preservation', () => {
  it('stores w:anchor="12345" as the STRING "12345" (not number 12345)', async () => {
    const hl = await loadFirstHyperlink(' w:anchor="12345"');
    expect(hl).toBeDefined();
    const anchor = hl?.getAnchor();
    expect(typeof anchor).toBe('string');
    expect(anchor).toBe('12345');
  });

  it('stores w:tooltip="42" as the STRING "42"', async () => {
    const hl = await loadFirstHyperlink(' w:anchor="target" w:tooltip="42"');
    const tooltip = hl?.getTooltip();
    expect(typeof tooltip).toBe('string');
    expect(tooltip).toBe('42');
  });

  it('stores w:tgtFrame="99" as the STRING "99"', async () => {
    const hl = await loadFirstHyperlink(' w:anchor="target" w:tgtFrame="99"');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const tgtFrame = (hl as any)?.tgtFrame;
    expect(typeof tgtFrame).toBe('string');
    expect(tgtFrame).toBe('99');
  });

  it('preserves non-numeric anchor (regression guard)', async () => {
    const hl = await loadFirstHyperlink(' w:anchor="_Toc123"');
    expect(hl?.getAnchor()).toBe('_Toc123');
  });

  it('string methods are callable on parsed numeric anchor', async () => {
    const hl = await loadFirstHyperlink(' w:anchor="12345"');
    const anchor = hl?.getAnchor() as string | undefined;
    // Pre-fix: anchor was the number 12345, .startsWith would throw.
    expect(() => anchor?.startsWith('1')).not.toThrow();
    expect(anchor?.startsWith('1')).toBe(true);
  });
});
