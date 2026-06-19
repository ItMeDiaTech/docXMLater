/**
 * setHeader()/setFooter() previously hardcoded the relationship target to
 * header1.xml/footer1.xml while HeaderFooterManager assigned the next
 * sequential part name. With a first/even-page header registered first, two
 * relationships targeted the same part and the default header's content sat
 * in an orphaned part — the relationship target must come from the
 * manager-assigned filename.
 */
import { Document } from '../../src/core/Document';
import { Header } from '../../src/elements/Header';
import { Footer } from '../../src/elements/Footer';

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

describe('C8: setHeader/setFooter relationships target the manager-assigned part', () => {
  it('gives each header/footer type a distinct part containing its own content', async () => {
    const doc = Document.create();
    let saved: Buffer;
    try {
      const first = Header.createFirst();
      first.createParagraph('FIRST PAGE HEADER');
      const even = Header.createEven();
      even.createParagraph('EVEN PAGE HEADER');
      const def = Header.createDefault();
      def.createParagraph('DEFAULT HEADER');

      // Register first/even before default — the order that previously made
      // the default relationship point at the first-page part
      doc.setFirstPageHeader(first);
      doc.setEvenPageHeader(even);
      doc.setHeader(def);

      const firstFooter = Footer.createFirst();
      firstFooter.createParagraph('FIRST PAGE FOOTER');
      const defFooter = Footer.createDefault();
      defFooter.createParagraph('DEFAULT FOOTER');

      doc.setFirstPageFooter(firstFooter);
      doc.setFooter(defFooter);

      doc.createParagraph('Body');
      saved = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = await JSZip.loadAsync(saved);
    const docXml = await zip.file('word/document.xml')!.async('string');
    const relsXml = await zip.file('word/_rels/document.xml.rels')!.async('string');

    const headerTargets: Record<string, string> = {};
    for (const type of ['first', 'even', 'default']) {
      const rId = getReferenceRId(docXml, 'headerReference', type);
      headerTargets[type] = getRelTarget(relsXml, rId);
    }
    expect(new Set(Object.values(headerTargets)).size).toBe(3);

    const footerTargets: Record<string, string> = {};
    for (const type of ['first', 'default']) {
      const rId = getReferenceRId(docXml, 'footerReference', type);
      footerTargets[type] = getRelTarget(relsXml, rId);
    }
    expect(new Set(Object.values(footerTargets)).size).toBe(2);

    // Each relationship resolves to a part holding that type's content
    const expectations: [string, string][] = [
      [headerTargets['first']!, 'FIRST PAGE HEADER'],
      [headerTargets['even']!, 'EVEN PAGE HEADER'],
      [headerTargets['default']!, 'DEFAULT HEADER'],
      [footerTargets['first']!, 'FIRST PAGE FOOTER'],
      [footerTargets['default']!, 'DEFAULT FOOTER'],
    ];
    for (const [target, marker] of expectations) {
      const partXml = await zip.file(`word/${target}`)!.async('string');
      expect(partXml).toContain(marker);
    }
  });
});
