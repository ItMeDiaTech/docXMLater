/**
 * `<w:pPrChange>` previous-indentation emission — `<w:ind>` mutual-exclusive
 * attribute handling.
 *
 * Per ECMA-376 Part 1 §17.3.1.12 CT_Ind, `w:hanging` and `w:firstLine` are
 * conceptually mutually exclusive — they represent opposite directions of
 * first-line indentation (hanging indent vs. first-line indent). Specifying
 * both produces ambiguous semantics, and Word itself always emits only one.
 *
 * The direct `<w:ind>` emitter in `Paragraph.generateParagraphPropertiesXml`
 * correctly uses if/else-if so only one is emitted (hanging wins when both
 * are set — the "stronger" semantic). The `<w:ind>` mirror inside
 * `<w:pPrChange>` previousProperties used independent `if`s and emitted
 * BOTH attributes, producing ambiguous tracked-change history.
 *
 * Same issue resolved inline for the twips pair; the hangingChars /
 * firstLineChars CJK pair already had if/else-if.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('<w:pPrChange> previous <w:ind> — mutual-exclusive attrs (§17.3.1.12)', () => {
  it('direct paragraph emission already prefers hanging when both set', () => {
    const p = new Paragraph();
    p.addText('x');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (p as any).formatting.indentation = { hanging: 360, firstLine: 720, left: 720 };
    const xml = XMLBuilder.elementToString(p.toXML());
    expect(xml).toMatch(/<w:ind[^>]*w:hanging="360"/);
    expect(xml).not.toMatch(/<w:ind[^>]*w:firstLine="720"/);
  });

  it('pPrChange previous-ind must also prefer hanging when both set', () => {
    const p = new Paragraph();
    p.addText('x');
    // Author a pPrChange whose previousProperties carry BOTH hanging and firstLine.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (p as any).formatting.pPrChange = {
      id: '1',
      author: 'Tester',
      date: '2026-01-01T00:00:00Z',
      previousProperties: {
        indentation: { hanging: 360, firstLine: 720, left: 720 },
      },
    };
    const xml = XMLBuilder.elementToString(p.toXML());
    // Extract the pPrChange block only.
    const changeBlock = xml.match(/<w:pPrChange[\s\S]*?<\/w:pPrChange>/)?.[0] ?? '';
    expect(changeBlock).toMatch(/<w:ind[^>]*w:hanging="360"/);
    // In the tracked-change previous state the firstLine must NOT appear
    // alongside hanging — they are mutually exclusive per §17.3.1.12.
    expect(changeBlock).not.toMatch(/<w:ind[^>]*w:firstLine="720"/);
  });

  it('pPrChange preserves firstLine when only firstLine is set (no hanging)', () => {
    const p = new Paragraph();
    p.addText('x');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (p as any).formatting.pPrChange = {
      id: '2',
      author: 'Tester',
      date: '2026-01-01T00:00:00Z',
      previousProperties: {
        indentation: { firstLine: 720, left: 720 },
      },
    };
    const xml = XMLBuilder.elementToString(p.toXML());
    const changeBlock = xml.match(/<w:pPrChange[\s\S]*?<\/w:pPrChange>/)?.[0] ?? '';
    expect(changeBlock).toMatch(/<w:ind[^>]*w:firstLine="720"/);
    expect(changeBlock).not.toMatch(/<w:ind[^>]*w:hanging/);
  });

  it('validator-clean round-trip for paragraph with pPrChange carrying both attrs', async () => {
    const doc = Document.create();
    const p = new Paragraph();
    p.addText('x');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (p as any).formatting.pPrChange = {
      id: '3',
      author: 'Tester',
      date: '2026-01-01T00:00:00Z',
      previousProperties: {
        indentation: { hanging: 360, firstLine: 720 },
      },
    };
    doc.addParagraph(p);
    await expect(doc.toBuffer()).resolves.toBeInstanceOf(Buffer);
    doc.dispose();
  });
});
