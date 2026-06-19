/**
 * TOC \t switch must be emitted once with all StyleName,Level doublets.
 *
 * Per ECMA-376 Part 1 §17.16.5.68, the \t switch takes a single
 * field-argument containing every custom style doublet
 * (\t "Heading 1,1,Heading 2,2"). Field switches are not repeatable;
 * Word honors only one \t instance, and DocumentParser likewise parses
 * only the first match, so emitting one switch per style silently drops
 * every style after the first.
 */

import { TableOfContents } from '../../src/elements/TableOfContents';

function countTSwitches(instruction: string): number {
  return (instruction.match(/\\t\s/g) ?? []).length;
}

describe('TOC \\t switch single field-argument (ECMA-376 §17.16.5.68)', () => {
  it('joins multiple includeStyles into one \\t switch', () => {
    const toc = new TableOfContents({
      includeStyles: [
        { styleName: 'Heading 1', level: 1 },
        { styleName: 'Heading 2', level: 2 },
        { styleName: 'AppendixHeading', level: 3 },
      ],
    });
    const instruction = toc.getFieldInstruction();
    expect(countTSwitches(instruction)).toBe(1);
    expect(instruction).toContain('\\t "Heading 1,1,Heading 2,2,AppendixHeading,3"');
  });

  it('emits a single doublet without a trailing comma for one style', () => {
    const toc = new TableOfContents({
      includeStyles: [{ styleName: 'MyStyle', level: 1 }],
    });
    const instruction = toc.getFieldInstruction();
    expect(countTSwitches(instruction)).toBe(1);
    expect(instruction).toMatch(/\\t "MyStyle,1"(?!,)/);
    expect(instruction).not.toContain('MyStyle,1,');
  });

  it('emits one \\t switch via createWithStyles factory', () => {
    const toc = TableOfContents.createWithStyles(['Style1', 'Style2']);
    const instruction = toc.getFieldInstruction();
    expect(countTSwitches(instruction)).toBe(1);
    expect(instruction).toContain('\\t "Style1,1,Style2,2"');
  });

  it('uses \\o (no \\t) when includeStyles is absent', () => {
    const toc = new TableOfContents({ levels: 3 });
    const instruction = toc.getFieldInstruction();
    expect(countTSwitches(instruction)).toBe(0);
    expect(instruction).toContain('\\o "1-3"');
  });
});
