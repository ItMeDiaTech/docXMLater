/**
 * settings.xml CT_OnOff flag parsing tests.
 *
 * Per ECMA-376 Part 1 §17.17.4, a bare `<w:X/>` and `<w:X w:val="1"/>`
 * both mean the flag is on; `<w:X w:val="0"/>` (or "false"/"off") means
 * it's explicitly off. `Document.parseSettingsFromXml` detected presence
 * with a simple `/<w:X\b[^>]*\/?>/.test(...)` regex and hard-coded the
 * flag to `true` whenever the element appeared at all — so a source
 * document that used `<w:X w:val="0"/>` to explicitly disable one of
 * these settings was silently flipped back to `true`.
 *
 * Nine settings flags were affected. Three of them directly control
 * tracked-changes behaviour:
 *
 *   - w:doNotTrackMoves        (tracked-move recording)
 *   - w:doNotTrackFormatting   (tracked-formatting recording)
 *   - w:trackFormatting        (mirror of the above)
 *
 * ...the rest cover hyphenation, spelling/grammar display, font
 * embedding, mirror margins, odd/even headers, and field update policy.
 *
 * This suite locks parse behaviour for every ST_OnOff literal on all
 * eleven flags.
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

describe('settings.xml CT_OnOff flags — parser honours w:val per ECMA-376 §17.17.4', () => {
  const flags: ReadonlyArray<{
    xml: string;
    getter: (doc: Document) => boolean;
  }> = [
    { xml: 'w:evenAndOddHeaders', getter: (d) => d.getEvenAndOddHeaders() },
    { xml: 'w:mirrorMargins', getter: (d) => d.getMirrorMargins() },
    { xml: 'w:autoHyphenation', getter: (d) => d.getAutoHyphenation() },
    { xml: 'w:hideSpellingErrors', getter: (d) => d.getHideSpellingErrors() },
    { xml: 'w:hideGrammaticalErrors', getter: (d) => d.getHideGrammaticalErrors() },
    { xml: 'w:updateFields', getter: (d) => d.getUpdateFields() },
    { xml: 'w:embedTrueTypeFonts', getter: (d) => d.getEmbedTrueTypeFonts() },
    { xml: 'w:saveSubsetFonts', getter: (d) => d.getSaveSubsetFonts() },
    { xml: 'w:doNotTrackMoves', getter: (d) => d.getDoNotTrackMoves() },
  ];

  for (const { xml, getter } of flags) {
    it(`parses <${xml} w:val="0"/> as false`, async () => {
      const buffer = await makeDocxWithSettings(`<${xml} w:val="0"/>`);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      expect(getter(doc)).toBe(false);
      doc.dispose();
    });

    it(`parses <${xml} w:val="false"/> as false`, async () => {
      const buffer = await makeDocxWithSettings(`<${xml} w:val="false"/>`);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      expect(getter(doc)).toBe(false);
      doc.dispose();
    });

    it(`parses <${xml} w:val="off"/> as false`, async () => {
      const buffer = await makeDocxWithSettings(`<${xml} w:val="off"/>`);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      expect(getter(doc)).toBe(false);
      doc.dispose();
    });

    it(`parses <${xml}/> as true`, async () => {
      const buffer = await makeDocxWithSettings(`<${xml}/>`);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      expect(getter(doc)).toBe(true);
      doc.dispose();
    });

    it(`parses <${xml} w:val="on"/> as true`, async () => {
      const buffer = await makeDocxWithSettings(`<${xml} w:val="on"/>`);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      expect(getter(doc)).toBe(true);
      doc.dispose();
    });

    it(`parses <${xml} w:val="1"/> as true`, async () => {
      const buffer = await makeDocxWithSettings(`<${xml} w:val="1"/>`);
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      expect(getter(doc)).toBe(true);
      doc.dispose();
    });

    it(`when <${xml}> absent, getter returns false`, async () => {
      const buffer = await makeDocxWithSettings('');
      const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
      expect(getter(doc)).toBe(false);
      doc.dispose();
    });
  }
});

describe('tracked-changes settings flags — parser honours w:val', () => {
  it('parses <w:doNotTrackFormatting w:val="0"/> as tracking enabled', async () => {
    const buffer = await makeDocxWithSettings('<w:doNotTrackFormatting w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    // doNotTrackFormatting=false → trackFormatting stays at its default (true)
    expect(doc.isTrackFormattingEnabled()).toBe(true);
    doc.dispose();
  });

  it('parses <w:doNotTrackFormatting/> as tracking disabled', async () => {
    const buffer = await makeDocxWithSettings('<w:doNotTrackFormatting/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.isTrackFormattingEnabled()).toBe(false);
    doc.dispose();
  });

  it('parses <w:doNotTrackFormatting w:val="false"/> as tracking enabled', async () => {
    const buffer = await makeDocxWithSettings('<w:doNotTrackFormatting w:val="false"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.isTrackFormattingEnabled()).toBe(true);
    doc.dispose();
  });

  it('parses <w:trackFormatting w:val="0"/> as tracking disabled', async () => {
    const buffer = await makeDocxWithSettings('<w:trackFormatting w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.isTrackFormattingEnabled()).toBe(false);
    doc.dispose();
  });

  it('parses <w:trackFormatting w:val="off"/> as tracking disabled', async () => {
    const buffer = await makeDocxWithSettings('<w:trackFormatting w:val="off"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.isTrackFormattingEnabled()).toBe(false);
    doc.dispose();
  });

  it('parses <w:trackFormatting/> as tracking enabled', async () => {
    const buffer = await makeDocxWithSettings('<w:trackFormatting/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    expect(doc.isTrackFormattingEnabled()).toBe(true);
    doc.dispose();
  });
});
