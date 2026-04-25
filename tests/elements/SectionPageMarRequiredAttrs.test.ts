/**
 * Section `<w:pgMar>` — required-attribute compliance.
 *
 * Per ECMA-376 Part 1 §17.6.11 CT_PageMar, ALL seven attributes are
 * declared with `use="required"`:
 *
 *   w:top, w:right, w:bottom, w:left, w:header, w:footer, w:gutter
 *
 * The current Section generator emits top/right/bottom/left unconditionally
 * (the `Margins` interface declares them as required TypeScript fields),
 * falls back to 720 twips for header/footer when unset, but only emits
 * `w:gutter` when it is explicitly defined — violating the schema's
 * required-attribute contract.
 *
 * Any `setMargins({ top, right, bottom, left })` call that omits gutter
 * produces `<w:pgMar ... />` missing the required `w:gutter` attribute.
 * Strict OOXML validators reject this with
 * *"The required attribute 'gutter' is missing"*.
 *
 * Fix: default `w:gutter` to `"0"` (the usual value — no gutter) so every
 * pgMar emission is schema-compliant.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Section } from '../../src/elements/Section';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('CT_PageMar required-attribute compliance (§17.6.11)', () => {
  it('emits w:gutter="0" when gutter not specified', () => {
    const section = new Section({
      margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
    });
    const xml = XMLBuilder.elementToString(section.toXML());
    // All 7 required attributes must be present.
    expect(xml).toMatch(/<w:pgMar[^>]*w:top="1440"/);
    expect(xml).toMatch(/<w:pgMar[^>]*w:right="1440"/);
    expect(xml).toMatch(/<w:pgMar[^>]*w:bottom="1440"/);
    expect(xml).toMatch(/<w:pgMar[^>]*w:left="1440"/);
    expect(xml).toMatch(/<w:pgMar[^>]*w:header="720"/);
    expect(xml).toMatch(/<w:pgMar[^>]*w:footer="720"/);
    expect(xml).toMatch(/<w:pgMar[^>]*w:gutter="0"/);
  });

  it('preserves explicit gutter when specified', () => {
    const section = new Section({
      margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, gutter: 720 },
    });
    const xml = XMLBuilder.elementToString(section.toXML());
    expect(xml).toMatch(/<w:pgMar[^>]*w:gutter="720"/);
  });

  it('a Document.create() document passes OOXML validator (pgMar has all required attrs)', async () => {
    const doc = Document.create();
    const p = new Paragraph();
    p.addText('test');
    doc.addParagraph(p);
    // Default margin from Document.create() doesn't set gutter — validator
    // would fail if our generator didn't supply the default.
    await expect(doc.toBuffer()).resolves.toBeInstanceOf(Buffer);
    doc.dispose();
  });
});
