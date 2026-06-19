/**
 * Revision.createPropertiesElement() must emit schema-valid OOXML local names
 * for non-run property-change snapshots: API-style keys (alignment) map to
 * their schema elements (w:jc per CT_PPrBase), object-valued properties
 * (spacing, indentation, borders, table widths) are serialized instead of
 * silently dropped, and children follow the xsd:sequence order of the
 * containing property element.
 */
import { Revision } from '../../src/elements/Revision';
import { Run } from '../../src/elements/Run';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

function serialize(revision: Revision): string {
  const xml = revision.toXML();
  expect(xml).not.toBeNull();
  return XMLBuilder.elementToString(xml!);
}

describe('Property-change snapshots use OOXML element names', () => {
  it('maps alignment to w:jc inside w:pPrChange/w:pPr', () => {
    const revision = Revision.createParagraphPropertiesChange('Editor', new Run('text'), {
      alignment: 'left',
    });

    const serialized = serialize(revision);
    expect(serialized).toContain('<w:jc w:val="left"');
    expect(serialized).not.toContain('<w:alignment');
  });

  it('serializes object-valued spacing and indentation instead of dropping them', () => {
    const revision = Revision.createParagraphPropertiesChange('Editor', new Run('text'), {
      spacing: { before: 120, after: 240, line: 360, lineRule: 'auto' },
      indentation: { left: 720, hanging: 360 },
    });

    const serialized = serialize(revision);
    expect(serialized).toContain(
      '<w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"'
    );
    expect(serialized).toContain('<w:ind w:left="720" w:hanging="360"');
  });

  it('orders pPr children per the CT_PPrBase sequence', () => {
    const revision = Revision.createParagraphPropertiesChange('Editor', new Run('text'), {
      alignment: 'center',
      keepNext: true,
      spacing: { before: 100 },
    });

    const serialized = serialize(revision);
    const keepNextIdx = serialized.indexOf('<w:keepNext');
    const spacingIdx = serialized.indexOf('<w:spacing');
    const jcIdx = serialized.indexOf('<w:jc');
    expect(keepNextIdx).toBeGreaterThan(-1);
    expect(spacingIdx).toBeGreaterThan(keepNextIdx);
    expect(jcIdx).toBeGreaterThan(spacingIdx);
  });

  it('keeps explicit false as w:val="0" per CT_OnOff tri-state', () => {
    const revision = Revision.createParagraphPropertiesChange('Editor', new Run('text'), {
      keepNext: false,
    });

    const serialized = serialize(revision);
    expect(serialized).toContain('<w:keepNext w:val="0"');
  });

  it('serializes numbering snapshots as w:numPr with w:ilvl/w:numId', () => {
    const revision = Revision.createParagraphPropertiesChange('Editor', new Run('text'), {
      numbering: { numId: 5, level: 1 },
    });

    const serialized = serialize(revision);
    expect(serialized).toContain('<w:numPr><w:ilvl w:val="1"');
    expect(serialized).toContain('<w:numId w:val="5"');
  });

  it('serializes table width snapshots with w:w/w:type', () => {
    const revision = Revision.createTablePropertiesChange('Editor', new Run('text'), {
      tblW: { w: 5000, type: 'dxa' },
      alignment: 'center',
    });

    const serialized = serialize(revision);
    expect(serialized).toContain('<w:tblW w:w="5000" w:type="dxa"');
    expect(serialized).toContain('<w:jc w:val="center"');
  });

  it('serializes border object snapshots', () => {
    const revision = Revision.createTablePropertiesChange('Editor', new Run('text'), {
      borders: { top: { style: 'single', size: 4, color: 'FF0000' } },
    });

    const serialized = serialize(revision);
    expect(serialized).toMatch(/<w:tblBorders><w:top w:val="single" w:sz="4" w:color="FF0000"/);
  });
});
