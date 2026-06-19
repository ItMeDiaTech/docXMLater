/**
 * setText on an empty (isEmpty) hyperlink must clear the flag so the text is
 * serialized; clone() of an empty hyperlink must retain the self-closing form.
 *
 * An empty/invisible hyperlink serializes as a childless <w:hyperlink/>. Before
 * the fix, setText() updated the in-memory text but never cleared _isEmpty, so
 * toXML() kept emitting an empty element and the assigned text silently
 * vanished from the saved document. clone() likewise dropped the isEmpty flag,
 * so a clone of an empty hyperlink serialized with a run child — diverging from
 * the snapshot it represents in tracked-change deletion pairs.
 */

import { Hyperlink } from '../../src/elements/Hyperlink';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('empty hyperlink setText / clone', () => {
  it('setText on an isEmpty hyperlink emits a run child', () => {
    const link = new Hyperlink({
      url: 'https://example.com',
      relationshipId: 'rId1',
      isEmpty: true,
    });

    // Sanity: starts self-closing (no children)
    expect(link.toXML().children ?? []).toHaveLength(0);

    link.setText('Visible text');

    const xmlEl = link.toXML();
    expect((xmlEl.children ?? []).length).toBeGreaterThan(0);

    const xml = XMLBuilder.elementToString(xmlEl);
    expect(xml).toContain('<w:r>');
    expect(xml).toContain('Visible text');
    expect(link.getText()).toBe('Visible text');
  });

  it('clone of an isEmpty hyperlink stays self-closing', () => {
    const link = new Hyperlink({
      url: 'https://example.com',
      relationshipId: 'rId1',
      isEmpty: true,
    });

    const copy = link.clone();
    expect(copy.isEmpty()).toBe(true);
    expect(copy.toXML().children ?? []).toHaveLength(0);
  });

  it('clone of a non-empty hyperlink keeps its run children', () => {
    const link = Hyperlink.createExternal('https://example.com', 'Hello');
    link.setRelationshipId('rId1');

    const copy = link.clone();
    expect(copy.isEmpty()).toBe(false);
    expect((copy.toXML().children ?? []).length).toBeGreaterThan(0);
  });
});
