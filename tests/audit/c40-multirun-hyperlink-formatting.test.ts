/**
 * Word splits a hyperlink's display text into one w:r per formatting change,
 * so a single-run Hyperlink model flattens multi-run links on round-trip:
 * the text of every run was concatenated into one run that carried only the
 * first run's rPr, silently dropping bold/italic/color on later runs.
 *
 * These tests pin the multi-run model: every parsed run survives with its
 * own rPr, both through toXML() and through a full load -> save round trip.
 */
import { Document } from '../../src/core/Document';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { Run } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

const MULTI_RUN_LINK_BODY =
  '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">plain </w:t></w:r>' +
  '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/></w:rPr><w:t>BOLDPART</w:t></w:r>';

/** Builds a docx whose only hyperlink has two runs with differing rPr. */
async function buildDocWithMultiRunHyperlink(): Promise<Buffer> {
  const doc = Document.create();
  const para = doc.createParagraph();
  para.addHyperlink(Hyperlink.createExternal('https://example.com', 'placeholder'));
  const buffer = await doc.toBuffer();
  doc.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  const docXml = zip.getFileAsString('word/document.xml')!;
  const relId = /<w:hyperlink[^>]*r:id="(rId\d+)"/.exec(docXml)![1];
  const replaced = docXml.replace(
    /<w:hyperlink[\s\S]*?<\/w:hyperlink>/,
    `<w:hyperlink r:id="${relId}" w:history="1">${MULTI_RUN_LINK_BODY}</w:hyperlink>`
  );
  zip.updateFile('word/document.xml', replaced);
  return zip.toBuffer();
}

function getHyperlinkRuns(xml: string): string[] {
  const link = /<w:hyperlink[^>]*>[\s\S]*?<\/w:hyperlink>/.exec(xml)?.[0] ?? '';
  return link.match(/<w:r[ >][\s\S]*?<\/w:r>/g) ?? [];
}

describe('C40: multi-run hyperlinks keep per-run formatting', () => {
  it('round-trips both runs with their own rPr (bold survives on run 2 only)', async () => {
    const input = await buildDocWithMultiRunHyperlink();

    const doc = await Document.loadFromBuffer(input);
    let saved: Buffer;
    try {
      saved = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = new ZipHandler();
    await zip.loadFromBuffer(saved);
    const savedXml = zip.getFileAsString('word/document.xml')!;

    const runs = getHyperlinkRuns(savedXml);
    expect(runs).toHaveLength(2);

    const plainRun = runs.find((r) => r.includes('plain '));
    const boldRun = runs.find((r) => r.includes('BOLDPART'));
    expect(plainRun).toBeDefined();
    expect(boldRun).toBeDefined();
    expect(boldRun).toMatch(/<w:b[ />]/);
    expect(plainRun).not.toMatch(/<w:b[ />]/);

    // Run order (and therefore display text) is preserved
    expect(savedXml.indexOf('plain ')).toBeLessThan(savedXml.indexOf('BOLDPART'));
  });

  it('exposes every parsed run on the loaded model', async () => {
    const input = await buildDocWithMultiRunHyperlink();

    const doc = await Document.loadFromBuffer(input);
    try {
      const links = doc.getHyperlinks();
      expect(links).toHaveLength(1);
      const link = links[0]!.hyperlink;
      expect(link.getText()).toBe('plain BOLDPART');

      const runs = link.getRuns();
      expect(runs).toHaveLength(2);
      expect(runs[0]!.getText()).toBe('plain ');
      expect(runs[0]!.getFormatting().bold).toBeUndefined();
      expect(runs[1]!.getText()).toBe('BOLDPART');
      expect(runs[1]!.getFormatting().bold).toBe(true);
    } finally {
      doc.dispose();
    }
  });

  it('toXML emits one w:r child per run set via setRuns', () => {
    const link = Hyperlink.createInternal('Section1', 'placeholder');
    link.setRuns([new Run('plain ', {}), new Run('BOLDPART', { bold: true })]);

    expect(link.getText()).toBe('plain BOLDPART');

    const serialized = XMLBuilder.elementToString(link.toXML());
    const runs = serialized.match(/<w:r[ >][\s\S]*?<\/w:r>/g) ?? [];
    expect(runs).toHaveLength(2);
    expect(runs[0]).toContain('plain ');
    expect(runs[0]).not.toMatch(/<w:b[ />]/);
    expect(runs[1]).toContain('BOLDPART');
    expect(runs[1]).toMatch(/<w:b[ />]/);
  });

  it('formatting setters apply to all runs without collapsing them', () => {
    const link = Hyperlink.createInternal('Section1', 'placeholder');
    link.setRuns([new Run('plain ', {}), new Run('BOLDPART', { bold: true })]);

    link.setColor('FF0000');

    const runs = link.getRuns();
    expect(runs).toHaveLength(2);
    expect(runs[0]!.getFormatting().color).toBe('FF0000');
    expect(runs[1]!.getFormatting().color).toBe('FF0000');
    // Per-run formatting the setter did not touch is preserved
    expect(runs[1]!.getFormatting().bold).toBe(true);
    expect(link.getText()).toBe('plain BOLDPART');
  });
});
