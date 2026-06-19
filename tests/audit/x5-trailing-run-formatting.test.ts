/**
 * replaceTextCrossRun() folded the unmatched tail of the last affected run
 * into the FIRST run and deleted the last run, so trailing text outside the
 * match silently took the first run's formatting (the last run's rPr was
 * discarded). The method's contract says partially consumed runs are
 * trimmed; these tests pin the in-place trim that keeps the tail's rPr.
 */
import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('X5: cross-run replace keeps trailing-run formatting', () => {
  it('trims the partially-consumed trailing run in place, keeping its rPr', () => {
    const para = new Paragraph();
    para.addRun(new Run('Hello {{na'));
    para.addRun(new Run('me}}, your balance is ', { bold: true, color: 'FF0000' }));

    const count = para.replaceTextCrossRun('{{name}}', 'Alice');

    expect(count).toBe(1);
    expect(para.getText()).toBe('Hello Alice, your balance is ');

    const runs = para.getRuns();
    expect(runs).toHaveLength(2);
    expect(runs[0]!.getText()).toBe('Hello Alice');
    expect(runs[0]!.getFormatting().bold).toBeUndefined();
    expect(runs[1]!.getText()).toBe(', your balance is ');
    expect(runs[1]!.getFormatting().bold).toBe(true);
    expect(runs[1]!.getFormatting().color).toBe('FF0000');
  });

  it('removes interior runs and keeps boundary-run formatting on both sides', () => {
    const para = new Paragraph();
    para.addRun(new Run('Dear {{', { italic: true }));
    para.addRun(new Run('name'));
    para.addRun(new Run('}}, welcome!', { bold: true }));

    const count = para.replaceTextCrossRun('{{name}}', 'Bob');

    expect(count).toBe(1);
    expect(para.getText()).toBe('Dear Bob, welcome!');

    const runs = para.getRuns();
    expect(runs).toHaveLength(2);
    expect(runs[0]!.getText()).toBe('Dear Bob');
    expect(runs[0]!.getFormatting().italic).toBe(true);
    expect(runs[1]!.getText()).toBe(', welcome!');
    expect(runs[1]!.getFormatting().bold).toBe(true);
  });

  it('still removes the last run when it is fully consumed', () => {
    const para = new Paragraph();
    para.addRun(new Run('Hel'));
    para.addRun(new Run('lo'));

    expect(para.replaceTextCrossRun('Hello', 'Hi')).toBe(1);
    expect(para.getRuns()).toHaveLength(1);
    expect(para.getText()).toBe('Hi');
  });

  it('handles multiple fragmented matches without corrupting offsets', () => {
    const para = new Paragraph();
    para.addRun(new Run('{{'));
    para.addRun(new Run('first'));
    para.addRun(new Run('}} and {{', { bold: true }));
    para.addRun(new Run('second'));
    para.addRun(new Run('}}'));

    expect(para.replaceTextCrossRun('{{first}}', 'A')).toBe(1);
    expect(para.replaceTextCrossRun('{{second}}', 'B')).toBe(1);
    expect(para.getText()).toBe('A and B');

    // The ' and ' tail came from the bold middle run and keeps its rPr
    const boldRun = para.getRuns().find((r) => r.getText().includes(' and '));
    expect(boldRun).toBeDefined();
    expect(boldRun!.getFormatting().bold).toBe(true);
  });

  it('fillTemplate preserves tail formatting in saved XML', async () => {
    const doc = Document.create();
    let saved: Buffer;
    try {
      const para = doc.createParagraph();
      para.addRun(new Run('Hello {{na'));
      para.addRun(new Run('me}}, your balance is positive', { bold: true }));

      const count = doc.fillTemplate({ name: 'Alice' });
      expect(count).toBe(1);

      saved = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const zip = new ZipHandler();
    await zip.loadFromBuffer(saved);
    const xml = zip.getFileAsString('word/document.xml')!;

    const runs = xml.match(/<w:r[ >][\s\S]*?<\/w:r>/g) ?? [];
    const tailRun = runs.find((r) => r.includes(', your balance is positive'));
    expect(tailRun).toBeDefined();
    expect(tailRun).toMatch(/<w:b[ />]/);

    const headRun = runs.find((r) => r.includes('Hello Alice'));
    expect(headRun).toBeDefined();
    expect(headRun).not.toMatch(/<w:b[ />]/);
  });
});
