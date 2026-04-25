/**
 * w:revisionView — w:markup and w:comments attribute round-trip.
 *
 * CT_TrackChangesView per ECMA-376 §17.15.1.77 exposes five ST_OnOff
 * attributes controlling what tracked-change markup is visible in the
 * reviewer pane:
 *
 *   w:markup          — all revision markup visible
 *   w:comments        — comment balloons visible
 *   w:insDel          — insertions/deletions visible         (ALREADY PARSED)
 *   w:formatting      — formatting changes visible           (ALREADY PARSED)
 *   w:inkAnnotations  — ink annotations visible              (ALREADY PARSED)
 *
 * Document.ts parsed only the last three — the first two were silently
 * stripped on round-trip. A document configuring the reviewer's default
 * markup view (e.g. hiding all revision markers initially) lost that
 * setting after any programmatic save.
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

describe('revisionView w:markup / w:comments (§17.15.1.77)', () => {
  it('parses w:markup="0" as showMarkup: false', async () => {
    const buffer = await makeDocxWithRevisionView('w:markup="0"');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const view = (doc as unknown as { revisionViewSettings?: { showMarkup?: boolean } })
      .revisionViewSettings;
    expect(view?.showMarkup).toBe(false);
    doc.dispose();
  });

  it('parses w:markup="true" as showMarkup: true', async () => {
    const buffer = await makeDocxWithRevisionView('w:markup="true"');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const view = (doc as unknown as { revisionViewSettings?: { showMarkup?: boolean } })
      .revisionViewSettings;
    expect(view?.showMarkup).toBe(true);
    doc.dispose();
  });

  it('parses w:comments="0" as showComments: false', async () => {
    const buffer = await makeDocxWithRevisionView('w:comments="0"');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const view = (doc as unknown as { revisionViewSettings?: { showComments?: boolean } })
      .revisionViewSettings;
    expect(view?.showComments).toBe(false);
    doc.dispose();
  });

  it('parses all five attributes together', async () => {
    const buffer = await makeDocxWithRevisionView(
      'w:markup="0" w:comments="0" w:insDel="1" w:formatting="1" w:inkAnnotations="0"'
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const view = (
      doc as unknown as {
        revisionViewSettings?: {
          showMarkup?: boolean;
          showComments?: boolean;
          showInsertionsAndDeletions?: boolean;
          showFormatting?: boolean;
          showInkAnnotations?: boolean;
        };
      }
    ).revisionViewSettings;
    expect(view?.showMarkup).toBe(false);
    expect(view?.showComments).toBe(false);
    expect(view?.showInsertionsAndDeletions).toBe(true);
    expect(view?.showFormatting).toBe(true);
    expect(view?.showInkAnnotations).toBe(false);
    doc.dispose();
  });

  it('round-trips w:markup and w:comments through Document save/load', async () => {
    const buffer = await makeDocxWithRevisionView('w:markup="0" w:comments="0"');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const rebuffered = await doc.toBuffer();
    doc.dispose();

    const zh = new ZipHandler();
    await zh.loadFromBuffer(rebuffered);
    const settingsXml = zh.getFileAsString('word/settings.xml') ?? '';
    expect(settingsXml).toMatch(/<w:revisionView\b[^>]*w:markup="0"/);
    expect(settingsXml).toMatch(/<w:revisionView\b[^>]*w:comments="0"/);
  });
});
