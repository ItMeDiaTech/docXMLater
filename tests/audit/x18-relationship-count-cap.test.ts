/**
 * RelationshipManager.fromXml previously threw CorruptedArchiveError for any
 * part with more than 1000 Relationship elements. Link- and image-heavy
 * documents (each unique hyperlink URL or image is one OPC relationship)
 * legitimately exceed 1000 entries while staying far under the 10MB size cap,
 * so they failed to load. The count cap is now a far-above-normal backstop
 * (>100000); the 10MB size cap is the primary protection.
 */
import { RelationshipManager } from '../../src/core/RelationshipManager';
import { CorruptedArchiveError } from '../../src/zip/errors';

function buildRelsXml(count: number): string {
  const parts: string[] = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
  ];
  for (let i = 1; i <= count; i++) {
    parts.push(
      `<Relationship Id="rId${i}" ` +
        `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ` +
        `Target="https://example.com/link/${i}" TargetMode="External"/>`
    );
  }
  parts.push('</Relationships>');
  return parts.join('');
}

describe('X18: RelationshipManager.fromXml relationship count cap', () => {
  it('parses a hyperlink-heavy part with more than 1000 relationships', () => {
    const count = 5000;
    const xml = buildRelsXml(count);

    // 5000 unique external hyperlinks is well under the 10MB size cap.
    expect(xml.length).toBeLessThan(10000000);

    const manager = RelationshipManager.fromXml(xml);

    expect(manager.getAllRelationships().length).toBe(count);
    expect(manager.getRelationship('rId1')?.getTarget()).toBe('https://example.com/link/1');
    expect(manager.getRelationship(`rId${count}`)?.getTarget()).toBe(
      `https://example.com/link/${count}`
    );
  });

  it('accepts exactly the backstop count (100000) but rejects one more', () => {
    // Pins the cap at its stated value: a future change that lowered it would
    // break the 100000 case; raising it would break the 100001 case.
    const minimal = (n: number) =>
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship/>'.repeat(n) +
      '</Relationships>';

    const atCap = minimal(100000);
    expect(atCap.length).toBeLessThan(10000000);
    expect(() => RelationshipManager.fromXml(atCap)).not.toThrow();

    expect(() => RelationshipManager.fromXml(minimal(100001))).toThrow(CorruptedArchiveError);
  });

  it('still rejects an absurd relationship count above the backstop', () => {
    // Minimal elements keep the part under the 10MB size cap so the count
    // backstop (not the size cap) is what rejects this input.
    const xml =
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship/>'.repeat(100001) +
      '</Relationships>';

    expect(xml.length).toBeLessThan(10000000);
    expect(() => RelationshipManager.fromXml(xml)).toThrow(CorruptedArchiveError);
  });
});
