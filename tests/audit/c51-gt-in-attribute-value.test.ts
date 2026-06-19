/**
 * Quote-aware tag-header scanning — a raw '>' inside an attribute value is
 * legal XML 1.0 (§2.3 AttValue requires escaping only '<', '&', and the
 * delimiting quote) and real serializers leave it unescaped.
 *
 * A bare indexOf('>') ended the tag header mid-attribute, truncating field
 * instructions like w:instr="IF 5 > 3" and injecting the attribute remainder
 * as garbage text that was re-serialized on save. extractElements and
 * extractBetweenTags shared the same quote-unaware scan.
 */

import { XMLParser } from '../../src/xml/XMLParser';

describe("XMLParser — raw '>' inside attribute values", () => {
  describe('parseToObject', () => {
    it('parses a fldSimple instruction containing a comparison operator', () => {
      const xml = '<w:fldSimple w:instr="IF 5 > 3"><w:r><w:t>yes</w:t></w:r></w:fldSimple>';
      const result: any = XMLParser.parseToObject(xml);

      expect(result['w:fldSimple']['@_w:instr']).toBe('IF 5 > 3');
      expect(result['w:fldSimple']['w:r']['w:t']).toBe('yes');
      // The attribute remainder must not leak into text content
      expect(result['w:fldSimple']['#text']).toBeUndefined();
    });

    it("handles '>' in single-quoted attribute values", () => {
      const xml = "<item cond='a > b'><sub/></item>";
      const result: any = XMLParser.parseToObject(xml);

      expect(result.item['@_cond']).toBe('a > b');
      expect(result.item.sub).toEqual({});
    });
  });

  describe('extractElements', () => {
    it("does not truncate an element whose attribute value contains '/>'", () => {
      const xml = '<w:p w:x="a/>b"><w:r><w:t>hi</w:t></w:r></w:p>';
      const elements = XMLParser.extractElements(xml, 'w:p');

      expect(elements).toHaveLength(1);
      expect(elements[0]).toBe(xml);
    });
  });

  describe('extractBetweenTags', () => {
    it("returns only inner content when the opening tag has '>' in an attribute", () => {
      const xml = '<w:pPr w:x="a > b"><w:jc w:val="center"/></w:pPr>';
      const inner = XMLParser.extractBetweenTags(xml, '<w:pPr', '</w:pPr>');

      expect(inner).toBe('<w:jc w:val="center"/>');
    });
  });
});
