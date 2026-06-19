/**
 * Revision.toXML() must keep the w:hyperlink wrapper (link target) when a
 * revision mixes runs and hyperlinks. Per ECMA-376, w:hyperlink is not a valid
 * child of CT_RunTrackChange, so the serialization splits into siblings: each
 * consecutive run group in its own w:ins/w:del and each hyperlink wrapping its
 * own w:ins/w:del. Dropping the wrapper silently downgraded the tracked link
 * to plain text and left its relationship unreferenced.
 */
import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { Revision } from '../../src/elements/Revision';
import { Run } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('Mixed-content revisions keep the hyperlink wrapper', () => {
  it('serializes run + hyperlink + run insertion as sibling w:ins segments', () => {
    const revision = Revision.createInsertion('Author', [
      new Run('before '),
      new Hyperlink({ url: 'https://example.com', text: 'Link', relationshipId: 'rId9' }),
      new Run(' after'),
    ]);
    revision.setId(3);

    const xml = revision.toXML();
    expect(xml).not.toBeNull();
    const serialized = XMLBuilder.elementToString(xml!);

    // The hyperlink wrapper with its link target survives
    expect(serialized).toContain('r:id="rId9"');
    expect(serialized).toMatch(/<w:hyperlink[^>]*r:id="rId9"[^>]*><w:ins/);

    // Runs before/after the link get their own w:ins siblings
    expect(serialized.startsWith('<w:ins')).toBe(true);
    expect(serialized).toContain('</w:ins><w:hyperlink');
    expect(serialized).toContain('</w:hyperlink><w:ins');

    // Content order is preserved across the segments
    const beforeIdx = serialized.indexOf('before ');
    const linkIdx = serialized.indexOf('<w:hyperlink');
    const afterIdx = serialized.indexOf(' after');
    expect(beforeIdx).toBeGreaterThan(-1);
    expect(linkIdx).toBeGreaterThan(beforeIdx);
    expect(afterIdx).toBeGreaterThan(linkIdx);
  });

  it('wraps mixed deletions: hyperlink-wrapped w:del uses w:delText', () => {
    const revision = Revision.createDeletion('Author', [
      new Run('cut '),
      new Hyperlink({ url: 'https://old.com', text: 'OldLink', relationshipId: 'rId4' }),
    ]);

    const xml = revision.toXML();
    expect(xml).not.toBeNull();
    const serialized = XMLBuilder.elementToString(xml!);

    expect(serialized).toMatch(/<w:hyperlink[^>]*r:id="rId4"[^>]*><w:del/);
    expect(serialized).toContain('OldLink');
    // Both the plain run and the hyperlink's runs are deleted text
    expect(serialized).not.toMatch(/<w:t[ >]/);
    expect((serialized.match(/<w:delText/g) || []).length).toBe(2);
  });

  it('keeps the tracked link and its relationship through save', async () => {
    const doc = Document.create();
    try {
      const para = new Paragraph();
      const revision = Revision.createInsertion('Author', [
        new Run('see '),
        new Hyperlink({ url: 'https://example.com/spec', text: 'the spec' }),
        new Run(' for details'),
      ]);
      para.addRevision(revision);
      doc.addParagraph(para);

      const buffer = await doc.toBuffer();
      const zip = new ZipHandler();
      await zip.loadFromBuffer(buffer);
      const docXml = zip.getFileAsString('word/document.xml')!;
      const rels = zip.getFileAsString('word/_rels/document.xml.rels')!;

      // Tracked-inserted link keeps its wrapper and r:id
      const match = docXml.match(/<w:hyperlink[^>]*r:id="(rId\d+)"[^>]*><w:ins/);
      expect(match).not.toBeNull();
      // The referenced relationship exists and targets the URL
      expect(rels).toContain(`Id="${match![1]}"`);
      expect(rels).toContain('https://example.com/spec');
      // Sibling runs stay tracked
      const seeIdx = docXml.indexOf('see ');
      const linkIdx = docXml.indexOf(match![0]);
      const detailsIdx = docXml.indexOf(' for details');
      expect(seeIdx).toBeGreaterThan(-1);
      expect(linkIdx).toBeGreaterThan(seeIdx);
      expect(detailsIdx).toBeGreaterThan(linkIdx);
    } finally {
      doc.dispose();
    }
  });
});
