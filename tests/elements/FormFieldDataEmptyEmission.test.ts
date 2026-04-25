/**
 * `<w:ffData>` emission — require at least one child per ECMA-376 §17.16.17.10.
 *
 * CT_FFData declares its content model as `<xsd:choice minOccurs="1"
 * maxOccurs="unbounded">`, which means AT LEAST ONE child element is
 * REQUIRED. Emitting `<w:ffData/>` with zero children produces
 * schema-invalid XML that strict OOXML validators reject with
 * *"The element has incomplete content. List of possible elements
 * expected: <name>, <label>, <tabIndex>, …"*.
 *
 * The previous Run emission unconditionally wrapped a `<w:fldChar
 * w:fldCharType="begin">` in a `<w:ffData>` whenever `formFieldData` was
 * truthy, even when that object was empty `{}`. The fix: fall back to a
 * bare `<w:fldChar>` (without ffData) when the form-field-data object
 * has no serializable fields.
 */

import { Run } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('<w:ffData> emission — CT_FFData minOccurs=1 child (§17.16.17.10)', () => {
  it('does NOT emit <w:ffData> when formFieldData is an empty object', () => {
    const run = new Run('');
    // Reach into the run's content directly — the public API doesn't
    // expose a way to build a malformed form-field, but the bug surfaces
    // through any caller that constructs a partial FormFieldData via
    // the parser or a third-party integration.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (run as any).content.push({
      type: 'fieldChar',
      fieldCharType: 'begin',
      formFieldData: {},
    });
    const xml = XMLBuilder.elementToString(run.toXML());
    // Must NOT emit <w:ffData/> with zero children — the validator rejects
    // this as "The element has incomplete content".
    expect(xml).not.toMatch(/<w:ffData\s*\/>/);
    expect(xml).not.toMatch(/<w:ffData[^>]*><\/w:ffData>/);
    // Should still emit the fldChar itself.
    expect(xml).toMatch(/<w:fldChar[^>]*w:fldCharType="begin"/);
  });

  it('still emits <w:ffData> when the form field has at least one real field', () => {
    const run = new Run('');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (run as any).content.push({
      type: 'fieldChar',
      fieldCharType: 'begin',
      formFieldData: { name: 'MyField' },
    });
    const xml = XMLBuilder.elementToString(run.toXML());
    expect(xml).toMatch(/<w:ffData>[\s\S]*<w:name\s+w:val="MyField"/);
  });
});
