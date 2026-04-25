/**
 * Section `<w:sectPr>` CT_OnOff children — explicit-false round-trip.
 *
 * `<w:formProt>`, `<w:noEndnote>`, `<w:titlePg>`, `<w:bidi>`, `<w:rtlGutter>`
 * are all OnOffType CT_OnOff children of CT_SectPrBase per ECMA-376 Part 1
 * §17.6.9 / §17.11.14 / §17.10.6 / §17.6.3 / §17.6.16.
 *
 * The parser already honors `w:val` and preserves the distinction between
 * absent (undefined), explicit-true (true), and explicit-false (false).
 * The generator historically only emitted for truthy values, so explicit
 * false collapsed into absent on round-trip — a fidelity loss.
 *
 * Per OOXML semantics at the section level, explicit false and absent are
 * semantically equivalent (there is no style inheritance for sections), so
 * this did not corrupt any document. But a load → save round-trip dropped
 * bytes the source authored — e.g. Word documents exported by third-party
 * tools that emit `<w:formProt w:val="0"/>` as an explicit "not protected"
 * marker lost that byte entirely.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithSectPrChildren(sectPrChildren: string): Promise<Buffer> {
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
    <w:p><w:r><w:t>hello</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      ${sectPrChildren}
    </w:sectPr>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

async function roundTripAndGetDocumentXml(buffer: Buffer): Promise<string> {
  const doc = await Document.loadFromBuffer(buffer);
  const out = await doc.toBuffer();
  doc.dispose();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(out);
  const docFile = zip.getFile('word/document.xml');
  const content = docFile?.content;
  return content instanceof Buffer ? content.toString('utf8') : String(content);
}

describe('sectPr CT_OnOff explicit-false round-trip', () => {
  describe('w:formProt (§17.6.9)', () => {
    it('round-trips true as <w:formProt/> or <w:formProt w:val="1"/>', async () => {
      const buffer = await makeDocxWithSectPrChildren('<w:formProt/>');
      const xml = await roundTripAndGetDocumentXml(buffer);
      // Any emission with val 1/true/on OR bare self-closing is acceptable.
      expect(xml).toMatch(/<w:formProt(?:\s+w:val="(?:1|true|on)")?\s*\/>/);
    });

    it('round-trips explicit false as <w:formProt w:val="0"/> (not dropped)', async () => {
      const buffer = await makeDocxWithSectPrChildren('<w:formProt w:val="0"/>');
      const xml = await roundTripAndGetDocumentXml(buffer);
      expect(xml).toMatch(/<w:formProt\s+w:val="(?:0|false|off)"\s*\/>/);
    });

    it('omits formProt when absent', async () => {
      const buffer = await makeDocxWithSectPrChildren('');
      const xml = await roundTripAndGetDocumentXml(buffer);
      expect(xml).not.toMatch(/<w:formProt\b/);
    });
  });

  describe('w:noEndnote (§17.11.14)', () => {
    it('round-trips true as <w:noEndnote/>', async () => {
      const buffer = await makeDocxWithSectPrChildren('<w:noEndnote/>');
      const xml = await roundTripAndGetDocumentXml(buffer);
      expect(xml).toMatch(/<w:noEndnote(?:\s+w:val="(?:1|true|on)")?\s*\/>/);
    });

    it('round-trips explicit false as <w:noEndnote w:val="0"/>', async () => {
      const buffer = await makeDocxWithSectPrChildren('<w:noEndnote w:val="0"/>');
      const xml = await roundTripAndGetDocumentXml(buffer);
      expect(xml).toMatch(/<w:noEndnote\s+w:val="(?:0|false|off)"\s*\/>/);
    });
  });

  describe('w:titlePg (§17.10.6)', () => {
    it('round-trips true as <w:titlePg/> or <w:titlePg w:val="1"/>', async () => {
      const buffer = await makeDocxWithSectPrChildren('<w:titlePg/>');
      const xml = await roundTripAndGetDocumentXml(buffer);
      expect(xml).toMatch(/<w:titlePg(?:\s+w:val="(?:1|true|on)")?\s*\/>/);
    });

    it('round-trips explicit false as <w:titlePg w:val="0"/>', async () => {
      const buffer = await makeDocxWithSectPrChildren('<w:titlePg w:val="0"/>');
      const xml = await roundTripAndGetDocumentXml(buffer);
      expect(xml).toMatch(/<w:titlePg\s+w:val="(?:0|false|off)"\s*\/>/);
    });
  });

  describe('w:bidi (§17.6.3 — section-level RTL)', () => {
    it('round-trips true as <w:bidi/>', async () => {
      const buffer = await makeDocxWithSectPrChildren('<w:bidi/>');
      const xml = await roundTripAndGetDocumentXml(buffer);
      expect(xml).toMatch(/<w:bidi(?:\s+w:val="(?:1|true|on)")?\s*\/>/);
    });

    it('round-trips explicit false as <w:bidi w:val="0"/>', async () => {
      const buffer = await makeDocxWithSectPrChildren('<w:bidi w:val="0"/>');
      const xml = await roundTripAndGetDocumentXml(buffer);
      expect(xml).toMatch(/<w:bidi\s+w:val="(?:0|false|off)"\s*\/>/);
    });
  });

  describe('w:rtlGutter (§17.6.16)', () => {
    it('round-trips true as <w:rtlGutter/>', async () => {
      const buffer = await makeDocxWithSectPrChildren('<w:rtlGutter/>');
      const xml = await roundTripAndGetDocumentXml(buffer);
      expect(xml).toMatch(/<w:rtlGutter(?:\s+w:val="(?:1|true|on)")?\s*\/>/);
    });

    it('round-trips explicit false as <w:rtlGutter w:val="0"/>', async () => {
      const buffer = await makeDocxWithSectPrChildren('<w:rtlGutter w:val="0"/>');
      const xml = await roundTripAndGetDocumentXml(buffer);
      expect(xml).toMatch(/<w:rtlGutter\s+w:val="(?:0|false|off)"\s*\/>/);
    });
  });
});
