/**
 * `<w:sdtPr><w:id w:val="0"/>` + `<w:tag w:val="123"/>` round-trip.
 *
 * Per ECMA-376 Part 1 §17.5.2.18 CT_SdtId, `w:val` is
 * ST_DecimalNumber — xsd:integer — so 0 is a legal SDT ID. Per
 * §17.5.2.34 CT_SdtText (w:tag), `w:val` is ST_String, accepting any
 * string including numeric-looking ones.
 *
 * Two bugs in the SDT-properties parser (`DocumentParser.ts:7634`):
 *   1. `<w:id w:val="0"/>` — XMLParser coerces `"0"` to the number
 *      `0`; the previous `if (idElement?.['@_w:val'])` truthy gate
 *      failed on 0 and silently dropped the SDT's identity. The
 *      emitter at `StructuredDocumentTag.ts:494` uses `!== undefined`,
 *      so the asymmetry broke round-trip for every SDT whose id was
 *      0.
 *   2. `<w:tag w:val="123"/>` (or other purely-numeric string) —
 *      XMLParser coerced `"123"` to the number `123`. The parser
 *      stored that number directly in `tag?: string`, violating the
 *      declared type. Subsequent code that used `tag.startsWith(…)`
 *      or similar string methods would either throw or silently
 *      misbehave.
 *
 * Iteration 112 fixes both:
 *   - `w:id`: gate on `isExplicitlySet` + `safeParseInt` + `!isNaN`.
 *   - `w:tag` / `w:alias`: gate on `!== undefined`, coerce to string.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { StructuredDocumentTag } from '../../src/elements/StructuredDocumentTag';

async function loadSdtProperties(sdtPrInner: string): Promise<{
  id?: number;
  tag?: string;
  alias?: string;
}> {
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
        ${sdtPrInner}
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
  const props = (sdt as any)?.properties as
    | { id?: number; tag?: string; alias?: string }
    | undefined;
  doc.dispose();
  return { id: props?.id, tag: props?.tag, alias: props?.alias };
}

describe('SDT <w:id>/<w:tag>/<w:alias> XMLParser-coercion round-trip', () => {
  it('preserves <w:id w:val="0"/> as the literal number 0 (previously dropped)', async () => {
    const { id } = await loadSdtProperties('<w:id w:val="0"/>');
    expect(id).toBe(0);
  });

  it('preserves <w:id w:val="12345"/> as number 12345 (regression guard)', async () => {
    const { id } = await loadSdtProperties('<w:id w:val="12345"/>');
    expect(id).toBe(12345);
  });

  it('preserves <w:tag w:val="123"/> as string "123" (previously stored as number)', async () => {
    const { tag } = await loadSdtProperties('<w:id w:val="1"/><w:tag w:val="123"/>');
    expect(typeof tag).toBe('string');
    expect(tag).toBe('123');
  });

  it('preserves <w:tag w:val="goog_rdk_xyz"/> (regression guard)', async () => {
    const { tag } = await loadSdtProperties('<w:id w:val="1"/><w:tag w:val="goog_rdk_xyz"/>');
    expect(tag).toBe('goog_rdk_xyz');
  });

  it('preserves <w:alias w:val="456"/> as string "456" (same coercion fix as tag)', async () => {
    const { alias } = await loadSdtProperties('<w:id w:val="1"/><w:alias w:val="456"/>');
    expect(typeof alias).toBe('string');
    expect(alias).toBe('456');
  });
});
