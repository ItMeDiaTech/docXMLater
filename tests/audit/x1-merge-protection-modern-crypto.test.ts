/**
 * w:documentProtection — merge path must preserve modern crypto attributes
 * and w:formatting.
 *
 * When settings.xml is mutated programmatically (e.g. enableTrackChanges())
 * before save, the original w:documentProtection element is stripped and
 * rebuilt from the parsed model. That rebuild must emit the same attribute
 * set as DocumentGenerator.generateSettings: the Word 2013+ crypto
 * attributes w:algorithmName / w:hashValue / w:saltValue (ISO/IEC 29500-4
 * §13) and w:formatting (ECMA-376 §17.15.1.29). Dropping them leaves
 * w:enforcement="1" with no password material, so protection can be
 * disabled without a password.
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

async function getSavedSettingsXml(rebuffered: Buffer): Promise<string> {
  const zh = new ZipHandler();
  await zh.loadFromBuffer(rebuffered);
  return zh.getFileAsString('word/settings.xml') ?? '';
}

describe('documentProtection merge path (settings modified before save)', () => {
  it('keeps w:algorithmName/w:hashValue/w:saltValue and w:formatting after enableTrackChanges()', async () => {
    const hashValue = 'aGFzaA==';
    const saltValue = 'c2FsdA==';
    const buffer = await makeDocxWithProtection(
      `w:edit="trackedChanges" w:enforcement="1" w:formatting="1" w:cryptProviderType="rsaAES" w:algorithmName="SHA-512" w:hashValue="${hashValue}" w:saltValue="${saltValue}"`
    );
    const doc = await Document.loadFromBuffer(buffer);
    try {
      doc.enableTrackChanges();
      const rebuffered = await doc.toBuffer();
      const settingsXml = await getSavedSettingsXml(rebuffered);

      expect(settingsXml).toMatch(/<w:documentProtection\b[^>]*w:edit="trackedChanges"/);
      expect(settingsXml).toContain('w:algorithmName="SHA-512"');
      expect(settingsXml).toContain(`w:hashValue="${hashValue}"`);
      expect(settingsXml).toContain(`w:saltValue="${saltValue}"`);
      expect(settingsXml).toMatch(/<w:documentProtection\b[^>]*w:formatting="1"/);
      // Legacy crypt attributes already emitted by the merge path must stay.
      expect(settingsXml).toContain('w:cryptProviderType="rsaAES"');
    } finally {
      doc.dispose();
    }
  });

  it('emits explicit w:formatting="0" after settings mutation (tri-state preserved)', async () => {
    const buffer = await makeDocxWithProtection(
      'w:edit="readOnly" w:enforcement="1" w:formatting="0"'
    );
    const doc = await Document.loadFromBuffer(buffer);
    try {
      doc.enableTrackChanges();
      const rebuffered = await doc.toBuffer();
      const settingsXml = await getSavedSettingsXml(rebuffered);
      expect(settingsXml).toMatch(/<w:documentProtection\b[^>]*w:formatting="0"/);
    } finally {
      doc.dispose();
    }
  });

  it('survives a second round-trip back into the parsed model', async () => {
    const hashValue = 'aGFzaA==';
    const saltValue = 'c2FsdA==';
    const buffer = await makeDocxWithProtection(
      `w:edit="trackedChanges" w:enforcement="1" w:formatting="1" w:algorithmName="SHA-512" w:hashValue="${hashValue}" w:saltValue="${saltValue}"`
    );
    const doc = await Document.loadFromBuffer(buffer);
    let rebuffered: Buffer;
    try {
      doc.enableTrackChanges();
      rebuffered = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const reloaded = await Document.loadFromBuffer(rebuffered);
    try {
      const prot = (
        reloaded as unknown as {
          documentProtection?: {
            formatting?: boolean;
            algorithmName?: string;
            hashValue?: string;
            saltValue?: string;
          };
        }
      ).documentProtection;
      expect(prot?.algorithmName).toBe('SHA-512');
      expect(prot?.hashValue).toBe(hashValue);
      expect(prot?.saltValue).toBe(saltValue);
      expect(prot?.formatting).toBe(true);
    } finally {
      reloaded.dispose();
    }
  });
});
