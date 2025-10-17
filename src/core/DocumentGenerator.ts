/**
 * DocumentGenerator - Handles XML generation for DOCX files
 * Converts structured data to OpenXML format
 */

import { XMLBuilder, XMLElement } from '../xml/XMLBuilder';
import { Paragraph } from '../elements/Paragraph';
import { Table } from '../elements/Table';
import { TableOfContentsElement } from '../elements/TableOfContentsElement';
import { Section } from '../elements/Section';
import { Hyperlink } from '../elements/Hyperlink';
import { ImageManager } from '../elements/ImageManager';
import { HeaderFooterManager } from '../elements/HeaderFooterManager';
import { CommentManager } from '../elements/CommentManager';
import { RelationshipManager } from './RelationshipManager';
import { DocumentProperties } from './Document';

/**
 * Body element types
 */
type BodyElement = Paragraph | Table | TableOfContentsElement;

/**
 * DocumentGenerator handles all XML generation logic
 */
export class DocumentGenerator {
  /**
   * Generates [Content_Types].xml
   */
  generateContentTypes(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
  }

  /**
   * Generates _rels/.rels
   */
  generateRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
  }

  /**
   * Generates word/document.xml with current body elements
   */
  generateDocumentXml(bodyElements: BodyElement[], section: Section): string {
    const bodyXmls: XMLElement[] = [];

    // Generate XML for each body element
    // Note: TableOfContentsElement.toXML() returns an array
    for (const element of bodyElements) {
      const xml = element.toXML();
      if (Array.isArray(xml)) {
        // TableOfContentsElement returns array of XMLElements
        bodyXmls.push(...xml);
      } else {
        // Paragraph and Table return single XMLElement
        bodyXmls.push(xml);
      }
    }

    // Add section properties at the end
    bodyXmls.push(section.toXML());
    return XMLBuilder.createDocument(bodyXmls);
  }

  /**
   * Generates docProps/core.xml
   */
  generateCoreProps(properties: DocumentProperties): string {
    const now = new Date();
    const created = properties.created || now;
    const modified = properties.modified || now;

    const formatDate = (date: Date): string => {
      return date.toISOString();
    };

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${XMLBuilder.escapeXmlText(properties.title || '')}</dc:title>
  <dc:subject>${XMLBuilder.escapeXmlText(properties.subject || '')}</dc:subject>
  <dc:creator>${XMLBuilder.escapeXmlText(properties.creator || 'DocXML')}</dc:creator>
  <cp:keywords>${XMLBuilder.escapeXmlText(properties.keywords || '')}</cp:keywords>
  <dc:description>${XMLBuilder.escapeXmlText(properties.description || '')}</dc:description>
  <cp:lastModifiedBy>${XMLBuilder.escapeXmlText(properties.lastModifiedBy || properties.creator || 'DocXML')}</cp:lastModifiedBy>
  <cp:revision>${properties.revision || 1}</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">${formatDate(created)}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${formatDate(modified)}</dcterms:modified>
</cp:coreProperties>`;
  }

  /**
   * Generates docProps/app.xml
   */
  generateAppProps(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>DocXML</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>0.1.0</AppVersion>
</Properties>`;
  }

  /**
   * Generates [Content_Types].xml with image extensions, headers/footers, and comments
   */
  generateContentTypesWithImagesHeadersFootersAndComments(
    imageManager: ImageManager,
    headerFooterManager: HeaderFooterManager,
    commentManager: CommentManager
  ): string {
    const images = imageManager.getAllImages();
    const headers = headerFooterManager.getAllHeaders();
    const footers = headerFooterManager.getAllFooters();
    const hasComments = commentManager.getCount() > 0;

    // Collect unique extensions
    const extensions = new Set<string>();
    for (const entry of images) {
      extensions.add(entry.image.getExtension());
    }

    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml +=
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n';

    // Default types
    xml +=
      '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n';
    xml += '  <Default Extension="xml" ContentType="application/xml"/>\n';

    // Image extensions
    for (const ext of extensions) {
      const mimeType = ImageManager.getMimeType(ext);
      xml += `  <Default Extension="${ext}" ContentType="${mimeType}"/>\n`;
    }

    // Override types
    xml +=
      '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n';
    xml +=
      '  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n';
    xml +=
      '  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>\n';

    // Header overrides
    for (const entry of headers) {
      xml += `  <Override PartName="/word/${entry.filename}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>\n`;
    }

    // Footer overrides
    for (const entry of footers) {
      xml += `  <Override PartName="/word/${entry.filename}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>\n`;
    }

    // Comments override
    if (hasComments) {
      xml +=
        '  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>\n';
    }

    xml +=
      '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>\n';
    xml +=
      '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\n';
    xml += '</Types>';

    return xml;
  }

  /**
   * Processes all hyperlinks in paragraphs and registers them with RelationshipManager
   */
  processHyperlinks(
    bodyElements: BodyElement[],
    headerFooterManager: HeaderFooterManager,
    relationshipManager: RelationshipManager
  ): void {
    // Get all paragraphs (from body and from headers/footers)
    const paragraphs = bodyElements.filter(
      (el): el is Paragraph => el instanceof Paragraph
    );

    // Also check headers and footers
    const headers = headerFooterManager.getAllHeaders();
    const footers = headerFooterManager.getAllFooters();

    for (const header of headers) {
      for (const element of header.header.getElements()) {
        if (element instanceof Paragraph) {
          this.processHyperlinksInParagraph(element, relationshipManager);
        }
      }
    }

    for (const footer of footers) {
      for (const element of footer.footer.getElements()) {
        if (element instanceof Paragraph) {
          this.processHyperlinksInParagraph(element, relationshipManager);
        }
      }
    }

    // Process body paragraphs
    for (const para of paragraphs) {
      this.processHyperlinksInParagraph(para, relationshipManager);
    }
  }

  /**
   * Processes hyperlinks in a single paragraph
   */
  private processHyperlinksInParagraph(
    paragraph: Paragraph,
    relationshipManager: RelationshipManager
  ): void {
    const content = paragraph.getContent();

    for (const item of content) {
      if (
        item instanceof Hyperlink &&
        item.isExternal() &&
        !item.getRelationshipId()
      ) {
        // Register external hyperlink with relationship manager
        const url = item.getUrl();
        if (url) {
          const relationship = relationshipManager.addHyperlink(url);
          item.setRelationshipId(relationship.getId());
        }
      }
    }
  }
}
