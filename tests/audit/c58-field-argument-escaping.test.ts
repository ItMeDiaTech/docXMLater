/**
 * Field-argument escaping (ECMA-376 §17.16.4.1).
 *
 * A literal double quote inside a quoted field-argument must be escaped
 * as \" and a literal backslash as \\ — otherwise an embedded quote
 * terminates the argument early and the remainder becomes stray switches
 * that Word misparses. The XML layer only escapes XML entities, so the
 * field-instruction factories must escape at field-code level.
 */

import {
  escapeFieldArgument,
  buildHyperlinkInstruction,
  createIFField,
} from '../../src/elements/FieldHelpers';
import { Field } from '../../src/elements/Field';

describe('escapeFieldArgument', () => {
  it('escapes embedded double quotes', () => {
    expect(escapeFieldArgument('See "Appendix A"')).toBe('See \\"Appendix A\\"');
  });

  it('escapes backslashes before quotes so the result is unambiguous', () => {
    expect(escapeFieldArgument('C:\\Docs\\file')).toBe('C:\\\\Docs\\\\file');
    expect(escapeFieldArgument('\\"')).toBe('\\\\\\"');
  });

  it('leaves plain text unchanged', () => {
    expect(escapeFieldArgument('plain text')).toBe('plain text');
  });
});

describe('buildHyperlinkInstruction escaping', () => {
  it('escapes quotes in the tooltip argument', () => {
    const instr = buildHyperlinkInstruction('https://example.com/', undefined, 'See "Appendix A"');

    expect(instr).toContain('\\o "See \\"Appendix A\\""');
  });

  it('escapes quotes and backslashes in url and anchor arguments', () => {
    const instr = buildHyperlinkInstruction('https://example.com/?q="x"', 'sec\\1');

    expect(instr).toContain('HYPERLINK "https://example.com/?q=\\"x\\""');
    expect(instr).toContain('\\l "sec\\\\1"');
  });
});

describe('Field factory escaping', () => {
  it('Field.createHyperlink escapes quotes in url and tooltip', () => {
    const field = Field.createHyperlink('https://example.com/"a"', 'Link', 'Tip "quoted"');

    expect(field.getInstruction()).toContain('HYPERLINK "https://example.com/\\"a\\""');
    expect(field.getInstruction()).toContain('\\o "Tip \\"quoted\\""');
  });

  it('Field.createTCEntry escapes quotes in the entry text', () => {
    const field = Field.createTCEntry('He said "hi"', 2);

    expect(field.getInstruction()).toBe('TC "He said \\"hi\\"" \\f C \\l 2');
  });

  it('Field.createXEEntry escapes quotes in entry and subentry', () => {
    const field = Field.createXEEntry('Main "M"', 'Sub "S"');

    expect(field.getInstruction()).toBe('XE "Main \\"M\\":Sub \\"S\\""');
  });
});

describe('createIFField escaping', () => {
  it('escapes quotes in true/false content', () => {
    const field = createIFField('Amount > 1000', 'High "Value"', 'Low \\ Normal');

    expect(field.getInstruction()).toBe(' IF Amount > 1000 "High \\"Value\\"" "Low \\\\ Normal" ');
  });
});
