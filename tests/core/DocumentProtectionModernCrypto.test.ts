/**
 * documentProtection — Word 2013+ crypto attributes round-trip.
 *
 * Per ISO/IEC 29500-4 §13 (transitional extensions to CT_DocProtect):
 *
 *   w:algorithmName — the strong algorithm identifier (e.g. "SHA-512").
 *                     Replaces the legacy `cryptAlgorithmSid` lookup-table
 *                     mapping used in Word 2003-2010.
 *   w:hashValue     — the password hash (base64).
 *   w:saltValue     — the password salt (base64).
 *
 * Word 2013+ always emits these three alongside (or instead of) the
 * legacy `hash` / `salt` / `cryptAlgorithmSid` attributes. Previously
 * none of the three were parsed or emitted, so a password-protected
 * document saved by modern Word dropped its password info on round-trip.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithProtection(protectionAttrs: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();
  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`
  );
  zipHandler.addFile(
    'word/settings.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:documentProtection ${protectionAttrs}/>
</w:settings>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>test</w:t></w:r></w:p></w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

type ProtAccess = {
  documentProtection?: {
    algorithmName?: string;
    hashValue?: string;
    saltValue?: string;
  };
};

describe('documentProtection modern crypto attributes', () => {
  it('parses w:algorithmName, w:hashValue, w:saltValue', async () => {
    // `AAAAAAAAAAAA=` and `BBBBBBBBBBBB=` are valid xsd:base64Binary
    // encodings (multiples-of-4 chars with correct padding).
    const buffer = await makeDocxWithProtection(
      'w:edit="readOnly" w:enforcement="1" w:algorithmName="SHA-512" w:hashValue="AAAAAAAAAAAA" w:saltValue="BBBBBBBBBBBB"'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const prot = (doc as unknown as ProtAccess).documentProtection;
    expect(prot?.algorithmName).toBe('SHA-512');
    expect(prot?.hashValue).toBe('AAAAAAAAAAAA');
    expect(prot?.saltValue).toBe('BBBBBBBBBBBB');
    doc.dispose();
  });

  it('leaves them undefined when absent', async () => {
    const buffer = await makeDocxWithProtection('w:edit="readOnly" w:enforcement="1"');
    const doc = await Document.loadFromBuffer(buffer);
    const prot = (doc as unknown as ProtAccess).documentProtection;
    expect(prot?.algorithmName).toBeUndefined();
    expect(prot?.hashValue).toBeUndefined();
    expect(prot?.saltValue).toBeUndefined();
    doc.dispose();
  });

  it('round-trips all three attributes through Document save/load', async () => {
    // Real valid base64 values — the OOXML validator enforces the
    // xsd:base64Binary constraint on w:hashValue / w:saltValue so bogus
    // strings like "abc==" fail schema validation.
    const hashValue = 'aGFzaA==';
    const saltValue = 'c2FsdA==';
    const buffer = await makeDocxWithProtection(
      `w:edit="readOnly" w:enforcement="1" w:algorithmName="SHA-512" w:hashValue="${hashValue}" w:saltValue="${saltValue}"`
    );
    const doc = await Document.loadFromBuffer(buffer);
    const rebuffered = await doc.toBuffer();
    doc.dispose();

    const zh = new ZipHandler();
    await zh.loadFromBuffer(rebuffered);
    const settingsXml = zh.getFileAsString('word/settings.xml') ?? '';
    expect(settingsXml).toContain('w:algorithmName="SHA-512"');
    expect(settingsXml).toContain(`w:hashValue="${hashValue}"`);
    expect(settingsXml).toContain(`w:saltValue="${saltValue}"`);
  });
});
