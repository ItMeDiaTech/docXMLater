/**
 * Auto-clean must keep an empty cleaned result instead of falling back to the
 * original markup-laden text.
 *
 * validateRunText sets cleanedText to the result of cleanXmlFromText whenever
 * XML patterns are detected and autoClean is on. For input that is entirely
 * XML markup (e.g. '<w:t></w:t>') the cleaned result is the empty string.
 * The previous `validation.cleanedText || text` reinstated the original markup
 * because '' is falsy — the exact case auto-clean exists for. The fix uses `??`
 * so only `undefined` (no cleaning performed) falls back to the original text.
 */

import { Run } from '../../src/elements/Run';
import { Hyperlink } from '../../src/elements/Hyperlink';

describe('Run/Hyperlink auto-clean keeps an empty cleaned result', () => {
  it('Run constructor strips fully-XML input to empty text', () => {
    const run = new Run('<w:t></w:t>');
    expect(run.getText()).toBe('');
  });

  it('Run constructor still cleans embedded markup from mixed input', () => {
    const run = new Run('hello <w:t>world</w:t>');
    expect(run.getText()).toBe('hello world');
  });

  it('Run.setText strips fully-XML input to empty text', () => {
    const run = new Run('placeholder');
    run.setText('<w:t></w:t>');
    expect(run.getText()).toBe('');
  });

  it('Hyperlink.setText strips fully-XML input to empty text', () => {
    const link = Hyperlink.createExternal('https://example.com', 'placeholder');
    link.setText('<w:t></w:t>');
    expect(link.getText()).toBe('');
  });
});
