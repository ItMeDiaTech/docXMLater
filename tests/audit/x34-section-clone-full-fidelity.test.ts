/**
 * Section.clone() must preserve every SectionProperties field, not just an
 * allow-list. The previous implementation built the cloned properties from an
 * explicit enumeration (pageSize, margins, columns, pageNumbering, headers,
 * footers, paperSource, docGrid, lineNumbering plus a few primitives) and
 * silently dropped pageBorders, footnotePr, endnotePr, noEndnote, formProt,
 * printerSettingsId, chapStyle and chapSep — losing page-border and section
 * note-numbering settings whenever a section was cloned (directly or via
 * DocumentContent.clone()).
 */

import { Section } from '../../src/elements/Section';

describe('X34: Section.clone() preserves all section properties', () => {
  function buildSection(): Section {
    return new Section({
      pageSize: { width: 12240, height: 15840, orientation: 'portrait' },
      pageBorders: {
        top: { style: 'single', size: 24, color: 'FF0000', space: 24 },
        bottom: { style: 'double', size: 18, color: '00FF00', space: 20 },
        left: { style: 'thick', size: 30, color: '0000FF', space: 16 },
        right: { style: 'dashed', size: 12, color: '123456', space: 12 },
        offsetFrom: 'page',
        display: 'allPages',
        zOrder: 'front',
      },
      footnotePr: { position: 'pageBottom', numberFormat: 'lowerRoman', startNumber: 3 },
      endnotePr: { position: 'sectEnd', numberFormat: 'decimal', startNumber: 1 },
      noEndnote: true,
      formProt: true,
      printerSettingsId: 'rId7',
      chapStyle: 2,
      chapSep: 'hyphen',
    });
  }

  it('carries the previously-dropped fields onto the clone', () => {
    const section = buildSection();
    const clonedProps = section.clone().getProperties();

    expect(clonedProps.pageBorders).toEqual({
      top: { style: 'single', size: 24, color: 'FF0000', space: 24 },
      bottom: { style: 'double', size: 18, color: '00FF00', space: 20 },
      left: { style: 'thick', size: 30, color: '0000FF', space: 16 },
      right: { style: 'dashed', size: 12, color: '123456', space: 12 },
      offsetFrom: 'page',
      display: 'allPages',
      zOrder: 'front',
    });
    expect(clonedProps.footnotePr).toEqual({
      position: 'pageBottom',
      numberFormat: 'lowerRoman',
      startNumber: 3,
    });
    expect(clonedProps.endnotePr).toEqual({
      position: 'sectEnd',
      numberFormat: 'decimal',
      startNumber: 1,
    });
    expect(clonedProps.noEndnote).toBe(true);
    expect(clonedProps.formProt).toBe(true);
    expect(clonedProps.printerSettingsId).toBe('rId7');
    expect(clonedProps.chapStyle).toBe(2);
    expect(clonedProps.chapSep).toBe('hyphen');
  });

  it('deep-copies nested pageBorders side objects so the clone is independent', () => {
    const section = buildSection();
    const cloned = section.clone();

    // Mutating the clone's nested border side must not affect the original.
    const clonedBorders = cloned.getProperties().pageBorders!;
    clonedBorders.top!.color = 'FFFFFF';
    clonedBorders.top!.size = 99;

    const originalBorders = section.getProperties().pageBorders!;
    expect(originalBorders.top!.color).toBe('FF0000');
    expect(originalBorders.top!.size).toBe(24);
  });

  it('deep-copies footnotePr/endnotePr so the clone is independent', () => {
    const section = buildSection();
    const cloned = section.clone();

    const clonedFootnote = cloned.getProperties().footnotePr!;
    clonedFootnote.startNumber = 42;

    expect(section.getProperties().footnotePr!.startNumber).toBe(3);
  });
});
