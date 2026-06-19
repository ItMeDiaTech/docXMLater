/**
 * Regression: StylesManager.cloneStyle must deep-copy nested formatting.
 *
 * The documented guarantee is a deep copy "including all run, paragraph, and
 * table formatting". A prior implementation spread source.getProperties() (a
 * shallow copy) into a new Style, so the clone and source aliased the same
 * runFormatting / paragraphFormatting / tableStyle objects. In-place mutators
 * (mergeWith, addConditionalFormatting, mutating getRunFormatting()'s return)
 * then corrupted the source through the shared references.
 */

import { StylesManager } from '../../src/formatting/StylesManager';
import { Style } from '../../src/formatting/Style';

describe('StylesManager.cloneStyle() deep-copies nested formatting', () => {
  it('does not alias the source tableStyle object', () => {
    const sm = StylesManager.create();

    const tableStyle = new Style({
      styleId: 'BaseTable',
      type: 'table',
      name: 'Base Table',
    });
    tableStyle.setRowBandSize(1);
    sm.addStyle(tableStyle);

    const clone = sm.cloneStyle('BaseTable', 'BaseTableAlt')!;

    // Adding conditional formatting to the clone must not leak into the source.
    clone.addConditionalFormatting({ type: 'firstRow' });

    const sourceTable = sm.getStyle('BaseTable')!.getProperties().tableStyle;
    const cloneTable = clone.getProperties().tableStyle;

    // The nested tableStyle objects must be distinct references.
    expect(cloneTable).not.toBe(sourceTable);
    expect(sourceTable?.conditionalFormatting).toBeUndefined();
    expect(cloneTable?.conditionalFormatting).toEqual([{ type: 'firstRow' }]);
  });

  it('does not alias the source runFormatting object', () => {
    const sm = StylesManager.create();

    const base = new Style({
      styleId: 'BaseChar',
      type: 'character',
      name: 'Base Char',
      runFormatting: { color: '000000', bold: true },
    });
    sm.addStyle(base);

    const clone = sm.cloneStyle('BaseChar', 'BaseCharAlt')!;

    const sourceRun = sm.getStyle('BaseChar')!.getRunFormatting();
    const cloneRun = clone.getRunFormatting();

    // The live run-formatting objects must be distinct references.
    expect(cloneRun).not.toBe(sourceRun);

    // Mutating the clone's live run formatting in place must not touch source.
    cloneRun!.color = 'FF0000';
    expect(sm.getStyle('BaseChar')!.getRunFormatting()?.color).toBe('000000');
  });

  it('mergeWith on the clone does not corrupt the source', () => {
    const sm = StylesManager.create();

    const base = new Style({
      styleId: 'MergeBase',
      type: 'paragraph',
      name: 'Merge Base',
      runFormatting: { color: '000000' },
      paragraphFormatting: { spacing: { before: 100, after: 100 } },
    });
    sm.addStyle(base);

    const clone = sm.cloneStyle('MergeBase', 'MergeBaseAlt')!;

    const override = new Style({
      styleId: 'Override',
      type: 'paragraph',
      name: 'Override',
      runFormatting: { color: 'FF0000', bold: true },
      paragraphFormatting: { spacing: { before: 999 } },
    });

    clone.mergeWith(override);

    // Source must retain its original values — mergeWith Object.assigns into the
    // clone's nested objects, which must not be shared with the source.
    const source = sm.getStyle('MergeBase')!;
    expect(source.getRunFormatting()?.color).toBe('000000');
    expect(source.getRunFormatting()?.bold).toBeUndefined();
    expect(source.getParagraphFormatting()?.spacing?.before).toBe(100);
  });
});
