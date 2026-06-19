/**
 * A pure load -> save round trip must not renumber header/footer parts.
 * Registration-order renumbering wrote each part's content to a different
 * part name than its relationship target, swapping header contents in
 * multi-header documents and orphaning single parts not named header1.xml.
 */
import { Document } from '../../src/core/Document';
import { Header } from '../../src/elements/Header';

const JSZip = require('jszip');

function getReferenceRId(docXml: string, tag: string, type: string): string {
  const match = new RegExp(`<w:${tag}[^>]*w:type="${type}"[^>]*/>`).exec(docXml);
  expect(match).not.toBeNull();
  const rId = /r:id="([^"]+)"/.exec(match![0]);
  expect(rId).not.toBeNull();
  return rId![1]!;
}

function getRelTarget(relsXml: string, rId: string): string {
  const match = new RegExp(`<Relationship[^>]*Id="${rId}"[^>]*/>`).exec(relsXml);
  expect(match).not.toBeNull();
  const target = /Target="([^"]+)"/.exec(match![0]);
  expect(target).not.toBeNull();
  return target![1]!;
}

/** Resolves a sectPr reference type to the content of the part it targets. */
async function resolveHeaderContent(zip: any, type: string): Promise<string> {
  const docXml = await zip.file('word/document.xml')!.async('string');
  const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string');
  const rId = getReferenceRId(docXml, 'headerReference', type);
  const target = getRelTarget(relsXml, rId);
  return zip.file(`word/${target}`)!.async('string');
}

describe('C7: load -> save round trip keeps header parts and relationships in sync', () => {
  it('does not swap contents of multi-header documents on a no-edit round trip', async () => {
    // titlePg document: first-page header registered before the default header,
    // the registration order that previously produced a content swap
    const doc1 = Document.create();
    const first = Header.createFirst();
    first.createParagraph('First Page Header Text');
    const def = Header.createDefault();
    def.createParagraph('Default Header Text');
    doc1.setFirstPageHeader(first);
    doc1.setHeader(def);
    doc1.createParagraph('Body');
    const buffer1 = await doc1.toBuffer();
    doc1.dispose();

    // Record the original part -> content mapping
    const zip1 = await JSZip.loadAsync(buffer1);
    const headerFiles1 = Object.keys(zip1.files)
      .filter((f: string) => /^word\/header\d+\.xml$/.exec(f))
      .sort();
    expect(headerFiles1).toHaveLength(2);
    const contents1 = new Map<string, string>();
    for (const file of headerFiles1) {
      contents1.set(file, await zip1.file(file)!.async('string'));
    }

    // Pure round trip, no edits
    const doc2 = await Document.loadFromBuffer(buffer1);
    let buffer2: Buffer;
    try {
      buffer2 = await doc2.toBuffer();
    } finally {
      doc2.dispose();
    }

    const zip2 = await JSZip.loadAsync(buffer2);

    // Same part names, each still holding its own content
    const headerFiles2 = Object.keys(zip2.files)
      .filter((f: string) => /^word\/header\d+\.xml$/.exec(f))
      .sort();
    expect(headerFiles2).toEqual(headerFiles1);
    for (const file of headerFiles1) {
      const original = contents1.get(file)!;
      const roundTripped = await zip2.file(file)!.async('string');
      const marker = original.includes('First Page Header Text')
        ? 'First Page Header Text'
        : 'Default Header Text';
      expect(roundTripped).toContain(marker);
    }

    // The sectPr references resolve to the right contents
    expect(await resolveHeaderContent(zip2, 'default')).toContain('Default Header Text');
    expect(await resolveHeaderContent(zip2, 'first')).toContain('First Page Header Text');
  });

  it('preserves a non-sequential part name on a no-edit round trip', async () => {
    // Build a docx whose only header part is word/header2.xml
    const doc1 = Document.create();
    const header = Header.createDefault();
    header.createParagraph('Original Header Text');
    doc1.setHeader(header);
    doc1.createParagraph('Body');
    const buf = await doc1.toBuffer();
    doc1.dispose();

    const zip = await JSZip.loadAsync(buf);
    const headerXml = await zip.file('word/header1.xml')!.async('string');
    zip.remove('word/header1.xml');
    zip.file('word/header2.xml', headerXml);
    let rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    rels = rels.replace('Target="header1.xml"', 'Target="header2.xml"');
    zip.file('word/_rels/document.xml.rels', rels);
    let contentTypes = await zip.file('[Content_Types].xml')!.async('string');
    contentTypes = contentTypes.replace('/word/header1.xml', '/word/header2.xml');
    zip.file('[Content_Types].xml', contentTypes);
    const renamed = await zip.generateAsync({ type: 'nodebuffer' });

    const doc2 = await Document.loadFromBuffer(renamed);
    let saved: Buffer;
    try {
      saved = await doc2.toBuffer();
    } finally {
      doc2.dispose();
    }

    const zip2 = await JSZip.loadAsync(saved);
    expect(zip2.file('word/header1.xml')).toBeNull();
    expect(await resolveHeaderContent(zip2, 'default')).toContain('Original Header Text');
  });
});
