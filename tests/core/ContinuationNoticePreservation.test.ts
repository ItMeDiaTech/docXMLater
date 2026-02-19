import { Document } from '../../src/core/Document';
import { FootnoteType } from '../../src/elements/Footnote';
import { EndnoteType } from '../../src/elements/Endnote';
import * as fs from 'fs';
import * as path from 'path';

describe('ContinuationNotice preservation', () => {
  it('should preserve continuationNotice after clearFootnotes/clearEndnotes', async () => {
    // Create a minimal DOCX with continuationNotice in footnotes/endnotes
    const doc = Document.create();
    const buf = await doc.toBuffer();
    doc.dispose();

    // Post-process: inject continuationNotice into footnotes.xml and endnotes.xml
    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(buf);

    const footnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationNotice" w:id="1"><w:p/></w:footnote>
</w:footnotes>`;

    const endnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
  <w:endnote w:type="continuationNotice" w:id="1"><w:p/></w:endnote>
</w:endnotes>`;

    zip.file('word/footnotes.xml', footnotesXml);
    zip.file('word/endnotes.xml', endnotesXml);

    const modifiedBuf = await zip.generateAsync({ type: 'nodebuffer' });

    // Load the document
    const doc2 = await Document.loadFromBuffer(modifiedBuf);

    // Verify continuationNotice was parsed
    const fm = doc2.getFootnoteManager();
    const em = doc2.getEndnoteManager();
    expect(fm.hasFootnote(1)).toBe(true);
    expect(em.hasEndnote(1)).toBe(true);

    // Simulate what Template_UI does: clear footnotes/endnotes
    doc2.clearFootnotes();
    doc2.clearEndnotes();

    // Verify continuationNotice survived clear()
    expect(fm.hasFootnote(1)).toBe(true);
    expect(em.hasEndnote(1)).toBe(true);
    expect(fm.getCount()).toBe(0); // No user footnotes
    expect(em.getCount()).toBe(0); // No user endnotes

    // Verify all special types are present
    const allFn = fm.getAllFootnotesWithSpecial();
    expect(allFn.length).toBe(3);
    expect(allFn.map(f => f.getType())).toEqual(
      expect.arrayContaining([
        FootnoteType.Separator,
        FootnoteType.ContinuationSeparator,
        FootnoteType.ContinuationNotice,
      ])
    );

    const allEn = em.getAllEndnotesWithSpecial();
    expect(allEn.length).toBe(3);
    expect(allEn.map(e => e.getType())).toEqual(
      expect.arrayContaining([
        EndnoteType.Separator,
        EndnoteType.ContinuationSeparator,
        EndnoteType.ContinuationNotice,
      ])
    );

    // Save and verify output contains continuationNotice
    const outBuf = await doc2.toBuffer();
    const outZip = await JSZip.loadAsync(outBuf);
    const outFootnotes = await outZip.file('word/footnotes.xml')?.async('string');
    const outEndnotes = await outZip.file('word/endnotes.xml')?.async('string');

    expect(outFootnotes).toContain('continuationNotice');
    expect(outFootnotes).toContain('w:id="1"');
    expect(outEndnotes).toContain('continuationNotice');
    expect(outEndnotes).toContain('w:id="1"');

    doc2.dispose();
  });

  it('should not allow removing continuationNotice via removeFootnote/removeEndnote', async () => {
    const doc = Document.create();
    const buf = await doc.toBuffer();
    doc.dispose();

    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(buf);

    zip.file('word/footnotes.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationNotice" w:id="1"><w:p/></w:footnote>
</w:footnotes>`);

    zip.file('word/endnotes.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
  <w:endnote w:type="continuationNotice" w:id="1"><w:p/></w:endnote>
</w:endnotes>`);

    const modifiedBuf = await zip.generateAsync({ type: 'nodebuffer' });
    const doc2 = await Document.loadFromBuffer(modifiedBuf);

    const fm = doc2.getFootnoteManager();
    const em = doc2.getEndnoteManager();

    // Attempt to remove special footnotes/endnotes
    expect(fm.removeFootnote(-1)).toBe(false);
    expect(fm.removeFootnote(0)).toBe(false);
    expect(fm.removeFootnote(1)).toBe(false); // continuationNotice is protected
    expect(em.removeEndnote(-1)).toBe(false);
    expect(em.removeEndnote(0)).toBe(false);
    expect(em.removeEndnote(1)).toBe(false);

    // They should all still exist
    expect(fm.hasFootnote(-1)).toBe(true);
    expect(fm.hasFootnote(0)).toBe(true);
    expect(fm.hasFootnote(1)).toBe(true);
    expect(em.hasEndnote(-1)).toBe(true);
    expect(em.hasEndnote(0)).toBe(true);
    expect(em.hasEndnote(1)).toBe(true);

    doc2.dispose();
  });

  it('should set nextId correctly after clear with continuationNotice', async () => {
    const doc = Document.create();
    const buf = await doc.toBuffer();
    doc.dispose();

    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(buf);

    zip.file('word/footnotes.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationNotice" w:id="1"><w:p/></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>User footnote</w:t></w:r></w:p></w:footnote>
</w:footnotes>`);

    const modifiedBuf = await zip.generateAsync({ type: 'nodebuffer' });
    const doc2 = await Document.loadFromBuffer(modifiedBuf);

    const fm = doc2.getFootnoteManager();

    // Before clear: 1 user footnote (id=2), continuationNotice (id=1) is special
    expect(fm.getCount()).toBe(1);
    expect(fm.hasFootnote(2)).toBe(true);

    doc2.clearFootnotes();

    // After clear: user footnote removed, continuationNotice preserved
    expect(fm.getCount()).toBe(0);
    expect(fm.hasFootnote(1)).toBe(true); // continuationNotice preserved
    expect(fm.hasFootnote(2)).toBe(false); // user footnote removed

    // nextId should skip past continuationNotice (id=1)
    expect(fm.getNextId()).toBe(2);

    // Creating a new footnote should get id=2
    const newFn = doc2.createFootnote('New footnote');
    expect(newFn.getId()).toBe(2);

    doc2.dispose();
  });

  it('should preserve continuationNotice through full round-trip with Original_60.docx', async () => {
    const filePath = path.resolve(__dirname, '../../Original_60.docx');
    if (!fs.existsSync(filePath)) {
      return; // Skip if file not available
    }

    const buf = fs.readFileSync(filePath);
    const doc = await Document.loadFromBuffer(buf);

    // Simulate Template_UI workflow: clear and recreate
    doc.clearFootnotes();
    doc.clearEndnotes();

    // Save
    const outBuf = await doc.toBuffer();

    // Verify output
    const JSZip = require('jszip');
    const outZip = await JSZip.loadAsync(outBuf);
    const outFootnotes = await outZip.file('word/footnotes.xml')?.async('string');
    const outEndnotes = await outZip.file('word/endnotes.xml')?.async('string');

    expect(outFootnotes).toContain('continuationNotice');
    expect(outFootnotes).toContain('w:id="1"');
    expect(outEndnotes).toContain('continuationNotice');
    expect(outEndnotes).toContain('w:id="1"');

    // Also verify settings.xml references are consistent
    const settingsXml = await outZip.file('word/settings.xml')?.async('string');
    if (settingsXml) {
      // If settings references footnote id=1, footnotes.xml must have it
      if (settingsXml.includes('<w:footnote w:id="1"')) {
        expect(outFootnotes).toContain('w:id="1"');
      }
      if (settingsXml.includes('<w:endnote w:id="1"')) {
        expect(outEndnotes).toContain('w:id="1"');
      }
    }

    doc.dispose();
  });
});
