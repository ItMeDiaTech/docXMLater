/**
 * `<wp:docPr>` `@name` and `@descr` — ST_String attributes must
 * round-trip as strings (type-contract safety).
 *
 * Per ECMA-376 Part 1 §20.4.2.5 CT_NonVisualDrawingProps, the docPr
 * element's `name` and `descr` attributes are both xsd:string. Any
 * string is valid including purely-numeric values.
 *
 * Bug: XMLParser's `parseAttributeValue: true` coerces purely-numeric
 * strings to JS numbers. The main image parser (DocumentParser.ts
 * §5920-5921) stored the raw coerced value:
 *
 *     name = docPrObj['@_name'] || 'image';        // could be number
 *     description = docPrObj['@_descr'] || 'Image'; // could be number
 *
 * `Image.name` / `Image.description` are declared `string`. Storing a
 * JS number violates the contract, and any downstream `.toLowerCase()`
 * / `.startsWith(...)` on the name or altText would throw. The
 * adjacent `@_title` parse (line 5923) already uses `String(...)`;
 * iteration 134 brings the sibling reads into line.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import fs from 'fs';
import path from 'path';

// Load a tiny real PNG from disk if available; otherwise generate a minimal one.
function makeMinimalPng(): Buffer {
  // Smallest valid 1x1 transparent PNG
  return Buffer.from(
    '89504E470D0A1A0A0000000D49484452000000010000000108060000001F15C4890000000D49444154789C62000100000500010D0A2DB40000000049454E44AE426082',
    'hex'
  );
}

async function loadDocWithImage(docPrAttrs: string) {
  const zipHandler = new ZipHandler();
  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
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
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`
  );
  zipHandler.addFile('word/media/image1.png', makeMinimalPng());
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="952500" cy="952500"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:docPr ${docPrAttrs}/>
            <wp:cNvGraphicFramePr/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr>
                    <pic:cNvPr id="0" name="Picture"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="rId2"/>
                    <a:stretch><a:fillRect/></a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="952500" cy="952500"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p><w:r><w:t>after</w:t></w:r></w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  return Document.loadFromBuffer(buffer);
}

function getFirstImage(doc: import('../../src/core/Document').Document) {
  const images = doc.getImages();
  return images[0]?.image;
}

describe('<wp:docPr> @name and @descr numeric type-contract', () => {
  it('stores @name="2010" as the STRING "2010"', async () => {
    const doc = await loadDocWithImage('id="1" name="2010" descr="MyImage"');
    const img = getFirstImage(doc);
    doc.dispose();
    expect(img).toBeDefined();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const name = (img as unknown as any).name;
    expect(typeof name).toBe('string');
    expect(name).toBe('2010');
  });

  it('stores @descr="42" as the STRING "42" (via getAltText())', async () => {
    const doc = await loadDocWithImage('id="1" name="Picture" descr="42"');
    const img = getFirstImage(doc);
    doc.dispose();
    const descr = img!.getAltText();
    expect(typeof descr).toBe('string');
    expect(descr).toBe('42');
  });

  it('string methods are callable on parsed numeric name', async () => {
    const doc = await loadDocWithImage('id="1" name="2010" descr="x"');
    const img = getFirstImage(doc);
    doc.dispose();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const name = (img as unknown as any).name as string;
    expect(() => name.startsWith('20')).not.toThrow();
    expect(name.startsWith('20')).toBe(true);
  });

  it('preserves non-numeric name "Chart 1" (regression guard)', async () => {
    const doc = await loadDocWithImage('id="1" name="Chart 1" descr="Quarterly chart"');
    const img = getFirstImage(doc);
    doc.dispose();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const name = (img as unknown as any).name;
    expect(name).toBe('Chart 1');
  });
});
