/**
 * TOC \p switch field-argument must be the literal separator character(s).
 *
 * Per ECMA-376 Part 1 §17.16.5.68, the \p switch argument is the literal
 * character sequence Word places between a TOC entry and its page number
 * (e.g. \p "-"). The dot/hyphen/underscore/none tokens are ST_TabTlc values
 * for tab-stop leaders and have no letter-code encoding in field
 * instructions; emitting \p "h" renders the letter 'h' as the visible
 * separator on every TOC line once Word updates the field.
 */

import { TableOfContents } from '../../src/elements/TableOfContents';

describe('TOC \\p switch literal separator (ECMA-376 §17.16.5.68)', () => {
  it('emits \\p "-" for tabLeader: hyphen', () => {
    const toc = new TableOfContents({ tabLeader: 'hyphen' });
    const instruction = toc.getFieldInstruction();
    expect(instruction).toContain('\\p "-"');
    expect(instruction).not.toContain('\\p "h"');
  });

  it('emits \\p "_" for tabLeader: underscore', () => {
    const toc = new TableOfContents({ tabLeader: 'underscore' });
    const instruction = toc.getFieldInstruction();
    expect(instruction).toContain('\\p "_"');
    expect(instruction).not.toContain('\\p "u"');
  });

  it('omits the \\p switch for tabLeader: none', () => {
    const toc = new TableOfContents({ tabLeader: 'none' });
    expect(toc.getFieldInstruction()).not.toContain('\\p');
  });

  it('omits the \\p switch for the default dot leader', () => {
    const toc = new TableOfContents();
    expect(toc.getFieldInstruction()).not.toContain('\\p');
  });

  it('emits the literal separator after configure()', () => {
    const toc = new TableOfContents();
    toc.configure({ tabLeader: 'hyphen' });
    expect(toc.getFieldInstruction()).toContain('\\p "-"');
  });
});
