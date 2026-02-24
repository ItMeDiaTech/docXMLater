/**
 * Tests for namespace ordering in XMLBuilder.createDocument()
 */

import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('Namespace Ordering', () => {
  it('should preserve document namespace order when namespaces are provided', () => {
    // Simulate a document with specific namespace order (as from original DOCX)
    const docNamespaces: Record<string, string> = {
      'xmlns:wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
      'xmlns:cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
      'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'xmlns:o': 'urn:schemas-microsoft-com:office:office',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
      'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    };

    const xml = XMLBuilder.createDocument([], docNamespaces);

    // The document namespaces should appear first in the output
    const wpcPos = xml.indexOf('xmlns:wpc=');
    const cxPos = xml.indexOf('xmlns:cx=');
    const mcPos = xml.indexOf('xmlns:mc=');

    // Document namespaces should come before framework defaults
    expect(wpcPos).toBeGreaterThan(-1);
    expect(cxPos).toBeGreaterThan(-1);

    // wpc should come before cx (preserving the order from docNamespaces)
    expect(wpcPos).toBeLessThan(cxPos);
    expect(cxPos).toBeLessThan(mcPos);
  });

  it('should include framework default namespaces not in document', () => {
    const docNamespaces: Record<string, string> = {
      'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    };

    const xml = XMLBuilder.createDocument([], docNamespaces);

    // Framework defaults that aren't in docNamespaces should still be present
    expect(xml).toContain('xmlns:r=');
    expect(xml).toContain('xmlns:wp=');
    expect(xml).toContain('xmlns:a=');
  });

  it('should use framework defaults when no document namespaces provided', () => {
    const xml = XMLBuilder.createDocument([]);

    // Should have standard framework namespaces
    expect(xml).toContain('xmlns:w=');
    expect(xml).toContain('xmlns:r=');
    expect(xml).toContain('xmlns:wp=');
  });

  it('should not duplicate namespaces from document that match framework defaults', () => {
    const docNamespaces: Record<string, string> = {
      'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    };

    const xml = XMLBuilder.createDocument([], docNamespaces);

    // Each namespace should appear exactly once
    const wCount = (xml.match(/xmlns:w="/g) || []).length;
    const rCount = (xml.match(/xmlns:r="/g) || []).length;
    expect(wCount).toBe(1);
    expect(rCount).toBe(1);
  });

  it('should preserve document namespace values when they differ from framework defaults', () => {
    const docNamespaces: Record<string, string> = {
      'xmlns:w': 'http://custom-namespace-for-testing',
    };

    const xml = XMLBuilder.createDocument([], docNamespaces);

    // Document value should win over framework default
    expect(xml).toContain('http://custom-namespace-for-testing');
  });
});
