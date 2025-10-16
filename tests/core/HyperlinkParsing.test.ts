/**
 * HyperlinkParsing.test.ts - Tests for hyperlink parsing when loading documents
 *
 * Tests the Document class's ability to parse <w:hyperlink> elements from existing DOCX files
 * and correctly reconstruct Hyperlink objects with their URLs, anchors, text, and formatting.
 */

import { Document } from '../../src/core/Document';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('Hyperlink Parsing', () => {
  describe('External Hyperlinks', () => {
    it('should parse external hyperlink with URL', async () => {
      // Create a mock document with an external hyperlink
      const mockDocument = await createMockDocument([
        {
          type: 'hyperlink',
          relationshipId: 'rId5',
          text: 'Visit our website',
          url: 'https://example.com',
        },
      ]);

      const doc = await Document.loadFromBuffer(mockDocument);
      const paragraphs = doc.getParagraphs();

      expect(paragraphs).toHaveLength(1);

      const para = paragraphs[0]!;
      const content = para.getContent();

      expect(content).toHaveLength(1);
      expect(content[0]).toBeInstanceOf(Hyperlink);

      const hyperlink = content[0] as Hyperlink;
      expect(hyperlink.getText()).toBe('Visit our website');
      expect(hyperlink.getUrl()).toBe('https://example.com');
      expect(hyperlink.isExternal()).toBe(true);
      expect(hyperlink.isInternal()).toBe(false);
    });

    it('should parse hyperlink with tooltip', async () => {
      const mockDocument = await createMockDocument([
        {
          type: 'hyperlink',
          relationshipId: 'rId5',
          text: 'Click here',
          url: 'https://example.com',
          tooltip: 'Visit our website',
        },
      ]);

      const doc = await Document.loadFromBuffer(mockDocument);
      const paragraphs = doc.getParagraphs();
      const hyperlink = paragraphs[0]!.getContent()[0] as Hyperlink;

      expect(hyperlink.getTooltip()).toBe('Visit our website');
    });

    it('should parse hyperlink with formatted text', async () => {
      const mockDocument = await createMockDocument([
        {
          type: 'hyperlink',
          relationshipId: 'rId5',
          text: 'Bold Link',
          url: 'https://example.com',
          formatting: {
            bold: true,
            color: '0563C1',
            underline: 'single',
          },
        },
      ]);

      const doc = await Document.loadFromBuffer(mockDocument);
      const paragraphs = doc.getParagraphs();
      const hyperlink = paragraphs[0]!.getContent()[0] as Hyperlink;

      const formatting = hyperlink.getFormatting();
      expect(formatting.bold).toBe(true);
      expect(formatting.color).toBe('0563C1');
      expect(formatting.underline).toBe('single');
    });

    it('should handle multiple hyperlinks in one paragraph', async () => {
      const mockDocument = await createMockDocument([
        {
          type: 'hyperlink',
          relationshipId: 'rId5',
          text: 'First link',
          url: 'https://first.com',
        },
        {
          type: 'hyperlink',
          relationshipId: 'rId6',
          text: 'Second link',
          url: 'https://second.com',
        },
      ]);

      const doc = await Document.loadFromBuffer(mockDocument);
      const paragraphs = doc.getParagraphs();
      const content = paragraphs[0]!.getContent();

      expect(content).toHaveLength(2);
      expect(content[0]).toBeInstanceOf(Hyperlink);
      expect(content[1]).toBeInstanceOf(Hyperlink);

      const link1 = content[0] as Hyperlink;
      const link2 = content[1] as Hyperlink;

      expect(link1.getText()).toBe('First link');
      expect(link1.getUrl()).toBe('https://first.com');
      expect(link2.getText()).toBe('Second link');
      expect(link2.getUrl()).toBe('https://second.com');
    });
  });

  describe('Internal Hyperlinks (Anchors)', () => {
    it('should parse internal hyperlink with anchor', async () => {
      const mockDocument = await createMockDocument([
        {
          type: 'hyperlink',
          anchor: 'Section1',
          text: 'Go to Section 1',
        },
      ]);

      const doc = await Document.loadFromBuffer(mockDocument);
      const paragraphs = doc.getParagraphs();
      const hyperlink = paragraphs[0]!.getContent()[0] as Hyperlink;

      expect(hyperlink.getText()).toBe('Go to Section 1');
      expect(hyperlink.getAnchor()).toBe('Section1');
      expect(hyperlink.isInternal()).toBe(true);
      expect(hyperlink.isExternal()).toBe(false);
    });

    it('should parse internal hyperlink with tooltip', async () => {
      const mockDocument = await createMockDocument([
        {
          type: 'hyperlink',
          anchor: 'Conclusion',
          text: 'Jump to conclusion',
          tooltip: 'Navigate to the conclusion section',
        },
      ]);

      const doc = await Document.loadFromBuffer(mockDocument);
      const paragraphs = doc.getParagraphs();
      const hyperlink = paragraphs[0]!.getContent()[0] as Hyperlink;

      expect(hyperlink.getAnchor()).toBe('Conclusion');
      expect(hyperlink.getTooltip()).toBe('Navigate to the conclusion section');
    });
  });

  describe('Edge Cases', () => {
    it('should handle hyperlink with empty text', async () => {
      const mockDocument = await createMockDocument([
        {
          type: 'hyperlink',
          relationshipId: 'rId5',
          text: '',
          url: 'https://example.com',
        },
      ]);

      const doc = await Document.loadFromBuffer(mockDocument);
      const paragraphs = doc.getParagraphs();
      const hyperlink = paragraphs[0]!.getContent()[0] as Hyperlink;

      // Should default to 'Link' when text is empty
      expect(hyperlink.getText()).toBe('Link');
    });

    it('should handle hyperlink with missing relationship', async () => {
      // Create hyperlink with relationship ID that doesn't exist in relationships
      const mockDocument = await createMockDocument([
        {
          type: 'hyperlink',
          relationshipId: 'rIdMissing',
          text: 'Broken link',
          url: undefined, // No relationship will be found
          skipRelationship: true,
        },
      ]);

      const doc = await Document.loadFromBuffer(mockDocument);
      const paragraphs = doc.getParagraphs();
      const hyperlink = paragraphs[0]!.getContent()[0] as Hyperlink;

      // Hyperlink should still be created but URL will be undefined
      expect(hyperlink.getText()).toBe('Broken link');
      expect(hyperlink.getUrl()).toBeUndefined();
      expect(hyperlink.getRelationshipId()).toBe('rIdMissing');
    });

    it('should handle hyperlink with special characters in text', async () => {
      const mockDocument = await createMockDocument([
        {
          type: 'hyperlink',
          relationshipId: 'rId5',
          text: 'Link & "Special" <Characters>',
          url: 'https://example.com',
        },
      ]);

      const doc = await Document.loadFromBuffer(mockDocument);
      const paragraphs = doc.getParagraphs();
      const hyperlink = paragraphs[0]!.getContent()[0] as Hyperlink;

      expect(hyperlink.getText()).toBe('Link & "Special" <Characters>');
    });
  });

  describe('Round-Trip Fidelity', () => {
    it('should preserve hyperlinks through load-save-load cycle', async () => {
      // Create document with hyperlink
      const doc1 = Document.create();
      const para1 = doc1.createParagraph();
      para1.addHyperlink(Hyperlink.createExternal('https://example.com', 'Test Link'));

      // Save to buffer
      const buffer1 = await doc1.toBuffer();

      // Load and verify
      const doc2 = await Document.loadFromBuffer(buffer1);
      const paragraphs2 = doc2.getParagraphs();
      const hyperlink2 = paragraphs2[0]!.getContent()[0] as Hyperlink;

      expect(hyperlink2.getText()).toBe('Test Link');
      expect(hyperlink2.getUrl()).toBe('https://example.com');

      // Save again
      const buffer2 = await doc2.toBuffer();

      // Load again and verify still intact
      const doc3 = await Document.loadFromBuffer(buffer2);
      const paragraphs3 = doc3.getParagraphs();
      const hyperlink3 = paragraphs3[0]!.getContent()[0] as Hyperlink;

      expect(hyperlink3.getText()).toBe('Test Link');
      expect(hyperlink3.getUrl()).toBe('https://example.com');
    });

    it('should preserve internal hyperlinks through round-trip', async () => {
      const doc1 = Document.create();
      const para1 = doc1.createParagraph();
      para1.addHyperlink(Hyperlink.createInternal('Section1', 'Go to Section 1'));

      const buffer1 = await doc1.toBuffer();
      const doc2 = await Document.loadFromBuffer(buffer1);
      const hyperlink2 = doc2.getParagraphs()[0]!.getContent()[0] as Hyperlink;

      expect(hyperlink2.getText()).toBe('Go to Section 1');
      expect(hyperlink2.getAnchor()).toBe('Section1');
      expect(hyperlink2.isInternal()).toBe(true);
    });
  });
});

/**
 * Helper function to create a mock DOCX document buffer with hyperlinks
 */
async function createMockDocument(
  hyperlinks: Array<{
    type: 'hyperlink';
    relationshipId?: string;
    anchor?: string;
    text: string;
    url?: string;
    tooltip?: string;
    formatting?: any;
    skipRelationship?: boolean;
  }>
): Promise<Buffer> {
  const zipHandler = new ZipHandler();

  // Create [Content_Types].xml
  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`
  );

  // Create _rels/.rels
  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`
  );

  // Generate hyperlink XML
  const hyperlinkElements = hyperlinks.map((link) => {
    const attrs: string[] = [];

    if (link.relationshipId) {
      attrs.push(`r:id="${link.relationshipId}"`);
    }
    if (link.anchor) {
      attrs.push(`w:anchor="${XMLBuilder.escapeXmlText(link.anchor)}"`);
    }
    if (link.tooltip) {
      attrs.push(`w:tooltip="${XMLBuilder.escapeXmlText(link.tooltip)}"`);
    }

    // Generate run with formatting
    const formattingXml = link.formatting
      ? `<w:rPr>
        ${link.formatting.bold ? '<w:b/>' : ''}
        ${link.formatting.italic ? '<w:i/>' : ''}
        ${link.formatting.underline ? `<w:u w:val="${link.formatting.underline}"/>` : ''}
        ${link.formatting.color ? `<w:color w:val="${link.formatting.color}"/>` : ''}
      </w:rPr>`
      : '';

    return `<w:hyperlink ${attrs.join(' ')}>
      <w:r>
        ${formattingXml}
        <w:t xml:space="preserve">${XMLBuilder.escapeXmlText(link.text)}</w:t>
      </w:r>
    </w:hyperlink>`;
  });

  // Create word/document.xml
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      ${hyperlinkElements.join('\n      ')}
    </w:p>
  </w:body>
</w:document>`
  );

  // Create word/_rels/document.xml.rels with hyperlink relationships
  const relationships = hyperlinks
    .filter((link) => link.relationshipId && !link.skipRelationship && link.url)
    .map(
      (link) =>
        `<Relationship Id="${link.relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${link.url}" TargetMode="External"/>`
    );

  zipHandler.addFile(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  ${relationships.join('\n  ')}
</Relationships>`
  );

  // Create minimal styles.xml
  zipHandler.addFile(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults/>
</w:styles>`
  );

  // Create minimal numbering.xml
  zipHandler.addFile(
    'word/numbering.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
`
  );

  // Create minimal docProps/core.xml
  zipHandler.addFile(
    'docProps/core.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:creator>Test</dc:creator>
</cp:coreProperties>`
  );

  // Create minimal docProps/app.xml
  zipHandler.addFile(
    'docProps/app.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Test</Application>
</Properties>`
  );

  return await zipHandler.toBuffer();
}
