/**
 * `parseHyperlinkFromObject` — bookmark `w:name` must round-trip as a
 * string (type-contract safety).
 *
 * Per ECMA-376 Part 1 §17.16.5 CT_Bookmark, `w:name` is `ST_String`.
 * XMLParser's `parseAttributeValue: true` coerces purely-numeric
 * bookmark names (e.g. `"12345"`) to JS numbers.
 *
 * Iteration 125 covered the hyperlink's own attributes (anchor,
 * tooltip, tgtFrame, docLocation) via the `toOptString` helper.
 * However, the sibling block that parses `<w:bookmarkStart>` nested
 * INSIDE `<w:hyperlink>` (§17.16.22 CT_Hyperlink permits EG_PContent,
 * so bookmarkStart can legally wrap runs inside a hyperlink) still
 * stored `name: bs['@_w:name']` without coercion:
 *
 *     const name = bs['@_w:name'];        // could be number
 *     const bookmark = new Bookmark({ name, ... });
 *
 * `Bookmark.name` is declared `string`. With `skipNormalization: true`
 * the raw value is stored directly, so a numeric name leaks through
 * to `bookmark.getName()` and breaks downstream string methods.
 *
 * Iteration 130 routes `name` through `String(...)` at the read site.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadDoc(paragraphBodyXml: string) {
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
    <w:p>${paragraphBodyXml}</w:p>
    <w:p><w:r><w:t>trailer</w:t></w:r></w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  return Document.loadFromBuffer(buffer);
}

describe('<w:hyperlink><w:bookmarkStart w:name="…"/> numeric type-contract', () => {
  it('stores hyperlink-wrapped <w:bookmarkStart w:name="12345"/> as STRING "12345"', async () => {
    const doc = await loadDoc(
      `<w:hyperlink w:anchor="target">
        <w:bookmarkStart w:id="1" w:name="12345"/>
        <w:r><w:t>link text</w:t></w:r>
       </w:hyperlink>`
    );
    const para = doc.getParagraphs()[0]!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const bookmarkStarts = (para as any).getBookmarksStart?.() ?? [];
    const bm = bookmarkStarts[0];
    doc.dispose();
    expect(bm).toBeDefined();
    const name = bm.getName();
    expect(typeof name).toBe('string');
    expect(name).toBe('12345');
  });

  it('stores <w:bookmarkStart w:name="42"/> as STRING "42"', async () => {
    const doc = await loadDoc(
      `<w:hyperlink w:anchor="target">
        <w:bookmarkStart w:id="2" w:name="42"/>
        <w:r><w:t>link</w:t></w:r>
       </w:hyperlink>`
    );
    const para = doc.getParagraphs()[0]!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const bookmarkStarts = (para as any).getBookmarksStart?.() ?? [];
    const name = bookmarkStarts[0].getName();
    doc.dispose();
    expect(typeof name).toBe('string');
    expect(name).toBe('42');
  });

  it('string methods are callable on parsed numeric bookmark name', async () => {
    const doc = await loadDoc(
      `<w:hyperlink w:anchor="target">
        <w:bookmarkStart w:id="3" w:name="12345"/>
        <w:r><w:t>link</w:t></w:r>
       </w:hyperlink>`
    );
    const para = doc.getParagraphs()[0]!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const bookmarkStarts = (para as any).getBookmarksStart?.() ?? [];
    const name = bookmarkStarts[0].getName() as string;
    doc.dispose();
    // Pre-fix: name was number 12345, .startsWith would throw.
    expect(() => name.startsWith('1')).not.toThrow();
    expect(name.startsWith('123')).toBe(true);
  });

  it('preserves non-numeric bookmark name "_Toc42" (regression guard)', async () => {
    const doc = await loadDoc(
      `<w:hyperlink w:anchor="target">
        <w:bookmarkStart w:id="4" w:name="_Toc42"/>
        <w:r><w:t>link</w:t></w:r>
       </w:hyperlink>`
    );
    const para = doc.getParagraphs()[0]!;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const bookmarkStarts = (para as any).getBookmarksStart?.() ?? [];
    const name = bookmarkStarts[0].getName();
    doc.dispose();
    expect(name).toBe('_Toc42');
  });
});
