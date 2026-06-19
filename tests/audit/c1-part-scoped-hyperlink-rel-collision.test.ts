/**
 * Per OPC (ECMA-376 Part 2), relationship IDs are part-scoped: rId1 in
 * word/_rels/header1.xml.rels is independent of rId1 in
 * word/_rels/document.xml.rels. Saving a loaded document whose header
 * contains an external hyperlink must not delete the unrelated main-rels
 * entry that happens to share the same ID (styles, settings, images...).
 */
import { Document } from '../../src/core/Document';
import { Header } from '../../src/elements/Header';
import { Paragraph } from '../../src/elements/Paragraph';
import { Relationship } from '../../src/core/Relationship';
import { Footnote, FootnoteType } from '../../src/elements/Footnote';
import { ZipHandler } from '../../src/zip/ZipHandler';

const HYPERLINK_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

const HEADER_WITH_HYPERLINK =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<w:p><w:hyperlink r:id="rId1"><w:r><w:t>Header Link</w:t></w:r></w:hyperlink></w:p>' +
  '</w:hdr>';

const HEADER_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  `<Relationship Id="rId1" Type="${HYPERLINK_REL_TYPE}" Target="https://example.com/header-link" TargetMode="External"/>` +
  '</Relationships>';

function getRelAttr(relsXml: string, rId: string, attr: string): string | undefined {
  const match = new RegExp(`<Relationship[^>]*Id="${rId}"[^>]*/>`).exec(relsXml);
  if (!match) return undefined;
  return new RegExp(`${attr}="([^"]+)"`).exec(match[0])?.[1];
}

/** Builds a docx whose header1.xml hyperlink uses the part-scoped ID rId1. */
async function buildDocWithHeaderHyperlink(): Promise<Buffer> {
  const doc = Document.create();
  const header = Header.createDefault();
  header.createParagraph('Placeholder');
  doc.setHeader(header);
  doc.createParagraph('Body');
  const buffer = await doc.toBuffer();
  doc.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  zip.updateFile('word/header1.xml', HEADER_WITH_HYPERLINK);
  zip.addFile('word/_rels/header1.xml.rels', HEADER_RELS);
  return zip.toBuffer();
}

describe('C1: part-scoped header hyperlink IDs must not delete main rels entries', () => {
  it('keeps the colliding document.xml.rels entry through a load -> save round trip', async () => {
    const input = await buildDocWithHeaderHyperlink();

    // Precondition: the main rels already use rId1 for a non-hyperlink part
    const zipBefore = new ZipHandler();
    await zipBefore.loadFromBuffer(input);
    const relsBefore = zipBefore.getFileAsString('word/_rels/document.xml.rels')!;
    const targetBefore = getRelAttr(relsBefore, 'rId1', 'Target');
    const typeBefore = getRelAttr(relsBefore, 'rId1', 'Type');
    expect(targetBefore).toBeDefined();
    expect(typeBefore).toBeDefined();
    expect(typeBefore).not.toBe(HYPERLINK_REL_TYPE);

    const doc = await Document.loadFromBuffer(input);
    let saved: Buffer;
    try {
      saved = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zipAfter = new ZipHandler();
    await zipAfter.loadFromBuffer(saved);
    const relsAfter = zipAfter.getFileAsString('word/_rels/document.xml.rels')!;
    // The colliding main entry survives with its original type and target
    expect(getRelAttr(relsAfter, 'rId1', 'Target')).toBe(targetBefore);
    expect(getRelAttr(relsAfter, 'rId1', 'Type')).toBe(typeBefore);

    // The header hyperlink stays part-scoped in header1.xml.rels
    const headerRels = zipAfter.getFileAsString('word/_rels/header1.xml.rels');
    expect(headerRels).toBeDefined();
    expect(headerRels).toContain('https://example.com/header-link');
    expect(headerRels).toContain(HYPERLINK_REL_TYPE);
  });

  it('still removes the main entry when it is genuinely the same hyperlink', async () => {
    const doc = Document.create();
    doc.createParagraph('Body text');

    const footnote = new Footnote({ id: 1, type: FootnoteType.Normal });
    const para = new Paragraph();
    const link = para.addHyperlink('https://example.com/footnote-link');
    link.setText('Footnote Link');
    link.setRelationshipId('rId600');
    footnote.addParagraph(para);
    doc.getFootnoteManager().register(footnote);

    doc.getRelationshipManager().addRelationship(
      Relationship.create({
        id: 'rId600',
        type: HYPERLINK_REL_TYPE,
        target: 'https://example.com/footnote-link',
        targetMode: 'External',
      })
    );

    const buffer = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(buffer);
    // Moved to the part rels, removed from the main rels
    const fnRels = zip.getFileAsString('word/_rels/footnotes.xml.rels');
    expect(fnRels).toContain('https://example.com/footnote-link');
    const docRels = zip.getFileAsString('word/_rels/document.xml.rels');
    expect(docRels).not.toContain('https://example.com/footnote-link');
  });
});
