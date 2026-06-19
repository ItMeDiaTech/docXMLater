/**
 * mergeTrackChangesIntoSettings must re-emit w:markup / w:comments.
 *
 * CT_TrackChangesView (ECMA-376 §17.15.1.77) carries five ST_OnOff attributes.
 * The parser populates showMarkup/showComments, and the new-document generator
 * emits them — but the settings *merge* path (taken whenever any settings change
 * marks _settingsModified) stripped the original <w:revisionView> and re-emitted
 * only insDel/formatting/inkAnnotations, dropping w:markup="0"/w:comments="0".
 * The fix extends both the emission gate and the attribute string in the merge
 * path to honour showMarkup/showComments.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithRevisionView(attrs: string): Promise<Buffer> {
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
  <w:revisionView ${attrs}/>
  <w:defaultTabStop w:val="720"/>
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

async function saveAndReadSettings(doc: Document): Promise<string> {
  const rebuffered = await doc.toBuffer();
  const zh = new ZipHandler();
  await zh.loadFromBuffer(rebuffered);
  return zh.getFileAsString('word/settings.xml') ?? '';
}

describe('revisionView w:markup / w:comments survive the settings merge path', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('keeps w:markup="0" and w:comments="0" after a settings mutation', async () => {
    const buffer = await makeDocxWithRevisionView('w:markup="0" w:comments="0"');
    doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    // Forces _settingsModified=true → mergeSettingsWithOriginal merge branch.
    doc.setDefaultTabStop(708);

    const settingsXml = await saveAndReadSettings(doc);
    expect(settingsXml).toMatch(/<w:revisionView\b[^>]*w:markup="0"/);
    expect(settingsXml).toMatch(/<w:revisionView\b[^>]*w:comments="0"/);
  });

  it('emits a revisionView element even when only markup/comments differ from defaults', async () => {
    // insDel/formatting/inkAnnotations all default true; only markup="0" set.
    const buffer = await makeDocxWithRevisionView('w:markup="0"');
    doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    doc.setDefaultTabStop(708);

    const settingsXml = await saveAndReadSettings(doc);
    expect(settingsXml).toMatch(/<w:revisionView\b[^>]*w:markup="0"/);
    expect(settingsXml).not.toMatch(/<w:revisionView\b[^>]*w:comments=/);
  });
});
