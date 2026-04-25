/**
 * w:documentProtection — `w:formatting` attribute round-trip.
 *
 * Per ECMA-376 Part 1 §17.15.1.29 (CT_DocProtect), `w:formatting` is a
 * ST_OnOff attribute that, when paired with `w:edit` protection, controls
 * whether formatting changes remain allowed even when edit restriction
 * is enforced. Directly relevant to the tracked-changes workflow: a
 * document with `w:edit="trackedChanges" w:formatting="true"` forces all
 * content edits to be tracked while still allowing formatting tweaks.
 *
 * Bug guarded against: neither the parser (Document.ts:~1033) nor the
 * generator (DocumentGenerator.ts:~989) handled `w:formatting` — so a
 * document setting it correctly was silently stripped on round-trip.
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

describe('documentProtection w:formatting (§17.15.1.29)', () => {
  it('parses w:formatting="true" as formatting: true', async () => {
    const buffer = await makeDocxWithProtection(
      'w:edit="trackedChanges" w:enforcement="1" w:formatting="true"'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const prot = (doc as unknown as { documentProtection?: { formatting?: boolean } })
      .documentProtection;
    expect(prot?.formatting).toBe(true);
    doc.dispose();
  });

  it('parses w:formatting="1" as formatting: true (ST_OnOff)', async () => {
    const buffer = await makeDocxWithProtection(
      'w:edit="readOnly" w:enforcement="1" w:formatting="1"'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const prot = (doc as unknown as { documentProtection?: { formatting?: boolean } })
      .documentProtection;
    expect(prot?.formatting).toBe(true);
    doc.dispose();
  });

  it('parses w:formatting="0" as formatting: false', async () => {
    const buffer = await makeDocxWithProtection(
      'w:edit="readOnly" w:enforcement="1" w:formatting="0"'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const prot = (doc as unknown as { documentProtection?: { formatting?: boolean } })
      .documentProtection;
    expect(prot?.formatting).toBe(false);
    doc.dispose();
  });

  it('leaves formatting undefined when attribute is absent', async () => {
    const buffer = await makeDocxWithProtection('w:edit="readOnly" w:enforcement="1"');
    const doc = await Document.loadFromBuffer(buffer);
    const prot = (doc as unknown as { documentProtection?: { formatting?: boolean } })
      .documentProtection;
    expect(prot?.formatting).toBeUndefined();
    doc.dispose();
  });

  it('round-trips w:formatting through Document save/load', async () => {
    const buffer = await makeDocxWithProtection(
      'w:edit="trackedChanges" w:enforcement="1" w:formatting="true"'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const rebuffered = await doc.toBuffer();
    doc.dispose();

    // Verify the re-saved settings.xml has w:formatting.
    const zh = new ZipHandler();
    await zh.loadFromBuffer(rebuffered);
    const settingsXml = zh.getFileAsString('word/settings.xml') ?? '';
    expect(settingsXml).toMatch(/<w:documentProtection\b[^>]*w:formatting="(1|true)"/);

    const reloaded = await Document.loadFromBuffer(rebuffered);
    const prot = (reloaded as unknown as { documentProtection?: { formatting?: boolean } })
      .documentProtection;
    expect(prot?.formatting).toBe(true);
    reloaded.dispose();
  });
});
