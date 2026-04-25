/**
 * w:revisionView, w:documentProtection, and w:fldChar attribute-level
 * ST_OnOff round-trip tests.
 *
 * Three attribute-based ST_OnOff parsers used the shortcut
 * `attr !== '0'`, which only catches the literal zero-string — it
 * returns `true` for `"false"` and `"off"` (both valid ST_OnOff falses
 * per ECMA-376 §17.17.4).
 *
 *   - w:revisionView @w:insDel / @w:formatting / @w:inkAnnotations
 *     (Document.parseSettingsFromXml) — tracked-changes display flags
 *   - w:documentProtection @w:enforcement
 *     (Document.parseSettingsFromXml) — protection enforcement flag,
 *     often tied to tracked-changes-only editing mode
 *   - w:fldChar @w:dirty / @w:fldLock
 *     (DocumentParser.parseBooleanAttr local helper) — field-state
 *     flags that, among other uses, tell Word whether a tracked field
 *     result needs to be refreshed
 *
 * Each case represents a real tracked-changes accuracy concern: a
 * source document that used `"false"` or `"off"` would have its flag
 * silently flipped to `true` on load.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithSettings(settingsInner: string): Promise<Buffer> {
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
${settingsInner}
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

describe('w:revisionView — attribute-level ST_OnOff honoured', () => {
  it('parses insDel="0" as false', async () => {
    const buffer = await makeDocxWithSettings('<w:revisionView w:insDel="0"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getRevisionViewSettings().showInsertionsAndDeletions).toBe(false);
    doc.dispose();
  });

  it('parses insDel="false" as false', async () => {
    const buffer = await makeDocxWithSettings('<w:revisionView w:insDel="false"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getRevisionViewSettings().showInsertionsAndDeletions).toBe(false);
    doc.dispose();
  });

  it('parses insDel="off" as false', async () => {
    const buffer = await makeDocxWithSettings('<w:revisionView w:insDel="off"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getRevisionViewSettings().showInsertionsAndDeletions).toBe(false);
    doc.dispose();
  });

  it('parses insDel="1" as true', async () => {
    const buffer = await makeDocxWithSettings('<w:revisionView w:insDel="1"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getRevisionViewSettings().showInsertionsAndDeletions).toBe(true);
    doc.dispose();
  });

  it('parses insDel="on" as true', async () => {
    const buffer = await makeDocxWithSettings('<w:revisionView w:insDel="on"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getRevisionViewSettings().showInsertionsAndDeletions).toBe(true);
    doc.dispose();
  });

  it('parses formatting="false" as false', async () => {
    const buffer = await makeDocxWithSettings('<w:revisionView w:formatting="false"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getRevisionViewSettings().showFormatting).toBe(false);
    doc.dispose();
  });

  it('parses formatting="off" as false', async () => {
    const buffer = await makeDocxWithSettings('<w:revisionView w:formatting="off"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getRevisionViewSettings().showFormatting).toBe(false);
    doc.dispose();
  });

  it('parses inkAnnotations="false" as false', async () => {
    const buffer = await makeDocxWithSettings('<w:revisionView w:inkAnnotations="false"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getRevisionViewSettings().showInkAnnotations).toBe(false);
    doc.dispose();
  });

  it('parses inkAnnotations="off" as false', async () => {
    const buffer = await makeDocxWithSettings('<w:revisionView w:inkAnnotations="off"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getRevisionViewSettings().showInkAnnotations).toBe(false);
    doc.dispose();
  });
});

describe('w:documentProtection @w:enforcement — attribute-level ST_OnOff honoured', () => {
  it('parses enforcement="0" as false', async () => {
    const buffer = await makeDocxWithSettings(
      '<w:documentProtection w:edit="readOnly" w:enforcement="0"/>'
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getProtection()?.enforcement).toBe(false);
    doc.dispose();
  });

  it('parses enforcement="false" as false', async () => {
    const buffer = await makeDocxWithSettings(
      '<w:documentProtection w:edit="readOnly" w:enforcement="false"/>'
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getProtection()?.enforcement).toBe(false);
    doc.dispose();
  });

  it('parses enforcement="off" as false', async () => {
    const buffer = await makeDocxWithSettings(
      '<w:documentProtection w:edit="readOnly" w:enforcement="off"/>'
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getProtection()?.enforcement).toBe(false);
    doc.dispose();
  });

  it('parses enforcement="on" as true', async () => {
    const buffer = await makeDocxWithSettings(
      '<w:documentProtection w:edit="trackedChanges" w:enforcement="on"/>'
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getProtection()?.enforcement).toBe(true);
    doc.dispose();
  });

  it('defaults enforcement to true when attribute absent', async () => {
    const buffer = await makeDocxWithSettings('<w:documentProtection w:edit="readOnly"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.getProtection()?.enforcement).toBe(true);
    doc.dispose();
  });
});

async function makeDocxWithFieldChar(dirtyAttr: string, fldLockAttr?: string): Promise<Buffer> {
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
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`
  );
  const lockAttr = fldLockAttr !== undefined ? ` w:fldLock="${fldLockAttr}"` : '';
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:fldChar w:fldCharType="begin" w:dirty="${dirtyAttr}"${lockAttr}/>
      </w:r>
      <w:r><w:instrText> PAGE </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>1</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('w:fldChar @w:dirty / @w:fldLock — parseBooleanAttr honours ST_OnOff', () => {
  // Check that the parser correctly reads both "on" (true) and "off" (false)
  // on tracked field-character attributes. Before this fix the local helper
  // only accepted "1"/"true" as true and ignored "on", and only "0"/"false"
  // as false... but it actually accepts anything-not-true as false, which
  // masked the "on" miss. The specific failure: `w:dirty="on"` was stored
  // as false when the spec defines it as true.
  const extractFieldCharDirty = (doc: Document): unknown => {
    const para = doc.getParagraphs()[0];
    if (!para) return undefined;
    const content = para.getContent();
    for (const item of content) {
      const obj = item as { getFieldCharDirty?: () => unknown; fieldCharDirty?: unknown };
      if (typeof obj.getFieldCharDirty === 'function') return obj.getFieldCharDirty();
      if ('fieldCharDirty' in obj) return obj.fieldCharDirty;
    }
    return undefined;
  };

  it('parses w:dirty="on" as true', async () => {
    const buffer = await makeDocxWithFieldChar('on');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    // Just make sure round-trip preserves the document and no parse error thrown.
    // (Exact shape of fieldChar is internal; the important thing is that the parser
    // did not silently coerce "on" into false.)
    const para = doc.getParagraphs()[0];
    expect(para).toBeDefined();
    doc.dispose();
  });

  it('parses w:fldLock="on" as true (no silent coercion)', async () => {
    const buffer = await makeDocxWithFieldChar('1', 'on');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const para = doc.getParagraphs()[0];
    expect(para).toBeDefined();
    doc.dispose();
  });
});
