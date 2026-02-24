import { Run } from '../../src/elements/Run';

describe('Run.clearMatchingFormatting', () => {
  it('should clear properties that match the style', () => {
    const run = new Run('Hello');
    run.setFont('Verdana');
    run.setSize(12);
    run.setColor('000000');

    run.clearMatchingFormatting({
      font: 'Verdana',
      size: 12,
      color: '000000',
    });

    const formatting = run.getFormatting();
    expect(formatting.font).toBeUndefined();
    expect(formatting.size).toBeUndefined();
    expect(formatting.color).toBeUndefined();
  });

  it('should preserve properties that differ from the style', () => {
    const run = new Run('Hello');
    run.setFont('Arial');
    run.setSize(14);
    run.setColor('FF0000');

    run.clearMatchingFormatting({
      font: 'Verdana',
      size: 12,
      color: '000000',
    });

    const formatting = run.getFormatting();
    expect(formatting.font).toBe('Arial');
    expect(formatting.size).toBe(14);
    expect(formatting.color).toBe('FF0000');
  });

  it('should preserve properties not defined in the style object', () => {
    const run = new Run('Hello');
    run.setFont('Verdana');
    run.setBold(true);
    run.setItalic(true);

    // Only clear font — bold and italic not in style object
    run.clearMatchingFormatting({ font: 'Verdana' });

    const formatting = run.getFormatting();
    expect(formatting.font).toBeUndefined();
    expect(formatting.bold).toBe(true);
    expect(formatting.italic).toBe(true);
  });

  it('should not clear anything when style object is empty', () => {
    const run = new Run('Hello');
    run.setFont('Verdana');
    run.setSize(12);

    run.clearMatchingFormatting({});

    const formatting = run.getFormatting();
    expect(formatting.font).toBe('Verdana');
    expect(formatting.size).toBe(12);
  });

  it('should return this for method chaining', () => {
    const run = new Run('Hello');
    run.setFont('Verdana');

    const result = run.clearMatchingFormatting({ font: 'Verdana' });
    expect(result).toBe(run);
  });
});
