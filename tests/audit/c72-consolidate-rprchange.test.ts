/**
 * consolidateRuns() merged adjacent runs whose current formatting matched by
 * rebuilding them via Run.createFromContent(), which carries only content +
 * formatting. Any w:rPrChange tracked-formatting history (ECMA-376
 * 17.13.5.32) on the merged runs was silently dropped from the model and the
 * serialized XML. rPrChange now acts as a merge boundary.
 */
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import type { RunFormatting } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

function runWithHistory(
  text: string,
  formatting: RunFormatting,
  id: number,
  author: string,
  previousProperties: Partial<RunFormatting>
): Run {
  const run = new Run(text, formatting);
  run.setPropertyChangeRevision({
    id,
    author,
    date: new Date('2024-03-01T00:00:00Z'),
    previousProperties,
  });
  return run;
}

describe('C72: consolidateRuns() preserves w:rPrChange history', () => {
  it('does not merge format-identical runs that both carry rPrChange', () => {
    const para = new Paragraph();
    para.addRun(runWithHistory('Hello ', { bold: true }, 1, 'Alice', { italic: true }));
    para.addRun(runWithHistory('World', { bold: true }, 2, 'Bob', { bold: false }));

    const eliminated = para.consolidateRuns();

    expect(eliminated).toBe(0);
    const runs = para.getRuns();
    expect(runs).toHaveLength(2);
    expect(runs[0]!.hasPropertyChangeRevision()).toBe(true);
    expect(runs[1]!.hasPropertyChangeRevision()).toBe(true);

    const xml = XMLBuilder.elementToString(para.toXML());
    expect((xml.match(/<w:rPrChange /g) ?? []).length).toBe(2);
    expect(xml).toContain('w:author="Alice"');
    expect(xml).toContain('w:author="Bob"');
  });

  it('does not merge when only one side carries rPrChange', () => {
    const para = new Paragraph();
    para.addRun(new Run('Hello ', { bold: true }));
    para.addRun(runWithHistory('World', { bold: true }, 3, 'Alice', { italic: true }));

    expect(para.consolidateRuns()).toBe(0);

    const runs = para.getRuns();
    expect(runs).toHaveLength(2);
    expect(runs[1]!.hasPropertyChangeRevision()).toBe(true);

    const xml = XMLBuilder.elementToString(para.toXML());
    expect((xml.match(/<w:rPrChange /g) ?? []).length).toBe(1);
  });

  it('still merges plain format-identical runs', () => {
    const para = new Paragraph();
    para.addRun(new Run('Hello ', { bold: true }));
    para.addRun(new Run('World', { bold: true }));

    expect(para.consolidateRuns()).toBe(1);
    expect(para.getRuns()).toHaveLength(1);
    expect(para.getText()).toBe('Hello World');
  });
});
