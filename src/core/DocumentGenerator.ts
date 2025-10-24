/**
 * DocumentGenerator - Handles XML generation for DOCX files
 * Converts structured data to OpenXML format
 */

import { XMLBuilder, XMLElement } from "../xml/XMLBuilder";
import { Paragraph } from "../elements/Paragraph";
import { Table } from "../elements/Table";
import { TableOfContentsElement } from "../elements/TableOfContentsElement";
import { StructuredDocumentTag } from "../elements/StructuredDocumentTag";
import { Section } from "../elements/Section";
import { Hyperlink } from "../elements/Hyperlink";
import { ImageManager } from "../elements/ImageManager";
import { HeaderFooterManager } from "../elements/HeaderFooterManager";
import { CommentManager } from "../elements/CommentManager";
import { FontManager } from "../elements/FontManager";
import { RelationshipManager } from "./RelationshipManager";
import { DocumentProperties } from "./Document";

/**
 * Body element types
 */
type BodyElement =
  | Paragraph
  | Table
  | TableOfContentsElement
  | StructuredDocumentTag;

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
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
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
  generateDocumentXml(
    bodyElements: BodyElement[],
    section: Section,
    namespaces: Record<string, string>
  ): string {
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
    return XMLBuilder.createDocument(bodyXmls, namespaces);
  }

  /**
   * Generates docProps/core.xml with extended properties
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
  <dc:title>${XMLBuilder.sanitizeXmlContent(properties.title || "")}</dc:title>
  <dc:subject>${XMLBuilder.sanitizeXmlContent(
    properties.subject || ""
  )}</dc:subject>
  <dc:creator>${XMLBuilder.sanitizeXmlContent(
    properties.creator || "DocXML"
  )}</dc:creator>
  <cp:keywords>${XMLBuilder.sanitizeXmlContent(
    properties.keywords || ""
  )}</cp:keywords>
  <dc:description>${XMLBuilder.sanitizeXmlContent(
    properties.description || ""
  )}</dc:description>
  <cp:lastModifiedBy>${XMLBuilder.sanitizeXmlContent(
    properties.lastModifiedBy || properties.creator || "DocXML"
  )}</cp:lastModifiedBy>
  <cp:revision>${properties.revision || 1}</cp:revision>${
    properties.category
      ? `\n  <cp:category>${XMLBuilder.sanitizeXmlContent(
          properties.category
        )}</cp:category>`
      : ""
  }${
    properties.contentStatus
      ? `\n  <cp:contentStatus>${XMLBuilder.sanitizeXmlContent(
          properties.contentStatus
        )}</cp:contentStatus>`
      : ""
  }${
    properties.language
      ? `\n  <dc:language>${XMLBuilder.sanitizeXmlContent(
          properties.language
        )}</dc:language>`
      : ""
  }
  <dcterms:created xsi:type="dcterms:W3CDTF">${formatDate(
    created
  )}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${formatDate(
    modified
  )}</dcterms:modified>
</cp:coreProperties>`;
  }

  /**
   * Generates docProps/app.xml with extended properties
   */
  generateAppProps(properties: DocumentProperties = {}): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>${XMLBuilder.sanitizeXmlContent(
    properties.application || "docxmlater"
  )}</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company>${XMLBuilder.sanitizeXmlContent(properties.company || "")}</Company>${
    properties.manager
      ? `\n  <Manager>${XMLBuilder.sanitizeXmlContent(
          properties.manager
        )}</Manager>`
      : ""
  }
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>${XMLBuilder.sanitizeXmlContent(
    properties.appVersion || properties.version || "1.0.0"
  )}</AppVersion>
</Properties>`;
  }

  /**
   * Generates docProps/custom.xml with custom properties
   */
  generateCustomProps(
    customProps: Record<string, string | number | boolean | Date>
  ): string {
    if (!customProps || Object.keys(customProps).length === 0) {
      return "";
    }

    const formatCustomValue = (
      key: string,
      value: string | number | boolean | Date,
      pid: number
    ): string => {
      if (typeof value === "string") {
        return `  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid}" name="${XMLBuilder.sanitizeXmlContent(
          key
        )}">
    <vt:lpwstr>${XMLBuilder.sanitizeXmlContent(value)}</vt:lpwstr>
  </property>`;
      } else if (typeof value === "number") {
        return `  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid}" name="${XMLBuilder.sanitizeXmlContent(
          key
        )}">
    <vt:r8>${value}</vt:r8>
  </property>`;
      } else if (typeof value === "boolean") {
        return `  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid}" name="${XMLBuilder.sanitizeXmlContent(
          key
        )}">
    <vt:bool>${value ? "true" : "false"}</vt:bool>
  </property>`;
      } else if (value instanceof Date) {
        return `  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid}" name="${XMLBuilder.sanitizeXmlContent(
          key
        )}">
    <vt:filetime>${value.toISOString()}</vt:filetime>
  </property>`;
      }
      return "";
    };

    const properties = Object.entries(customProps)
      .map(([key, value], index) => formatCustomValue(key, value, index + 2))
      .filter((prop) => prop !== "")
      .join("\n");

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
${properties}
</Properties>`;
  }

  /**
   * Generates [Content_Types].xml with image extensions, headers/footers, comments, and fonts
   * Preserves entries for files that exist in the loaded document (customXML, etc.)
   */
  generateContentTypesWithImagesHeadersFootersAndComments(
    imageManager: ImageManager,
    headerFooterManager: HeaderFooterManager,
    commentManager: CommentManager,
    zipHandler: any, // ZipHandler - to check file existence
    fontManager?: FontManager,
    hasCustomProperties: boolean = false
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

    // Font extensions (if FontManager provided)
    if (fontManager && fontManager.getCount() > 0) {
      const fontEntries = fontManager.generateContentTypeEntries();
      for (const entry of fontEntries) {
        xml += entry + "\n";
      }
    }

    // Check for embedded .ttf fonts from original document
    const files = zipHandler.getFilePaths?.() || [];
    const hasTtfFonts = files.some((f: string) => f.endsWith(".ttf"));
    if (hasTtfFonts) {
      xml +=
        '  <Default Extension="ttf" ContentType="application/x-font-ttf"/>\n';
    }

    // Override types
    xml +=
      '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n';
    xml +=
      '  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n';
    xml +=
      '  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>\n';

    // Required files (MUST be present for DOCX compliance)
    xml +=
      '  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>\n';
    xml +=
      '  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n';
    xml +=
      '  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>\n';

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

    // Only add app.xml if it exists (not all documents have it)
    if (zipHandler.hasFile?.("docProps/app.xml")) {
      xml +=
        '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>\n';
    }

    // Include custom properties if exists or will be created
    if (zipHandler.hasFile?.("docProps/custom.xml") || hasCustomProperties) {
      xml +=
        '  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>\n';
    }

    // Preserve customXML entries if they exist
    if (zipHandler.hasFile?.("customXML/item1.xml")) {
      xml +=
        '  <Override PartName="/customXML/item1.xml" ContentType="application/xml"/>\n';
    }
    if (zipHandler.hasFile?.("customXML/itemProps1.xml")) {
      xml +=
        '  <Override PartName="/customXML/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>\n';
    }

    xml += "</Types>";

    return xml;
  }

  /**
   * Clears ORPHANED hyperlink relationships from the RelationshipManager
   * Only removes relationships that don't have corresponding hyperlinks in the document
   *
   * This prevents corruption when paragraphs with hyperlinks are removed but
   * their relationships remain, causing Word's "unreadable content" error.
   * Preserves relationships for existing hyperlinks to maintain round-trip integrity.
   */
  private clearOrphanedHyperlinkRelationships(
    bodyElements: BodyElement[],
    headerFooterManager: HeaderFooterManager,
    relationshipManager: RelationshipManager
  ): void {
    // Step 1: Collect all relationship IDs currently used by hyperlinks
    const usedRelIds = new Set<string>();

    // Helper to scan paragraphs for hyperlink relationship IDs
    const scanParagraph = (para: Paragraph) => {
      for (const item of para.getContent()) {
        if (item instanceof Hyperlink && item.isExternal()) {
          const relId = item.getRelationshipId();
          if (relId) {
            usedRelIds.add(relId);
          }
        }
      }
    };

    // Helper to recursively scan any element type for hyperlinks
    const scanElement = (element: BodyElement | Paragraph | Table | StructuredDocumentTag): void => {
      if (element instanceof Paragraph) {
        // Scan paragraph content for hyperlinks
        scanParagraph(element);
      }
      else if (element instanceof Table) {
        // Scan all cells in the table
        for (let row = 0; row < element.getRowCount(); row++) {
          for (let col = 0; col < element.getColumnCount(); col++) {
            const cell = element.getCell(row, col);
            if (cell) {
              // Scan each paragraph in the cell
              const paragraphs = cell.getParagraphs();
              for (const para of paragraphs) {
                scanParagraph(para);
              }
            }
          }
        }
      }
      else if (element instanceof StructuredDocumentTag) {
        // Recursively scan SDT content (can contain Paragraphs, Tables, or nested SDTs)
        const content = element.getContent();
        for (const item of content) {
          scanElement(item); // Recursive call handles nested structures
        }
      }
      // TableOfContentsElement is for programmatic TOCs - real TOCs come as SDTs
    };

    // Scan body elements (handles all nested structures)
    for (const element of bodyElements) {
      scanElement(element);
    }

    // Scan headers (including tables and SDTs in headers)
    const headers = headerFooterManager.getAllHeaders();
    for (const header of headers) {
      for (const element of header.header.getElements()) {
        scanElement(element);
      }
    }

    // Scan footers (including tables and SDTs in footers)
    const footers = headerFooterManager.getAllFooters();
    for (const footer of footers) {
      for (const element of footer.footer.getElements()) {
        scanElement(element);
      }
    }

    // Step 2: Remove ONLY orphaned relationships (not used by any hyperlink)
    const allHyperlinkRels = relationshipManager.getRelationshipsByType(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    );

    for (const rel of allHyperlinkRels) {
      if (!usedRelIds.has(rel.getId())) {
        // This relationship is orphaned - remove it
        relationshipManager.removeRelationship(rel.getId());
      }
    }
  }

  /**
   * Processes all hyperlinks in paragraphs and registers them with RelationshipManager
   * Clears orphaned hyperlink relationships to prevent corruption while preserving valid ones
   */
  processHyperlinks(
    bodyElements: BodyElement[],
    headerFooterManager: HeaderFooterManager,
    relationshipManager: RelationshipManager
  ): void {
    // Clear ORPHANED hyperlink relationships to prevent corruption
    // This is critical when paragraphs are removed (e.g., via clearParagraphs())
    // but preserves relationships for existing hyperlinks (round-trip integrity)
    this.clearOrphanedHyperlinkRelationships(
      bodyElements,
      headerFooterManager,
      relationshipManager
    );

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
   *
   * **Validation:** Throws error if external hyperlink has no URL to prevent
   * document corruption per ECMA-376 §17.16.22.
   *
   * @throws {Error} If external hyperlink has undefined/empty URL
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

        // Validate that external hyperlink has a URL
        // This prevents invalid document generation and fails early with clear error
        if (!url) {
          throw new Error(
            `Invalid hyperlink in paragraph: External hyperlink "${item.getText()}" has no URL. ` +
              `This would create a corrupted document per ECMA-376 §17.16.22. ` +
              `Fix the hyperlink by providing a valid URL before saving.`
          );
        }

        const relationship = relationshipManager.addHyperlink(url);
        item.setRelationshipId(relationship.getId());
      }
    }
  }

  /**
   * Generates word/fontTable.xml
   * Required for DOCX compliance - defines fonts used in the document
   */
  generateFontTable(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="Calibri">
    <w:panose1 w:val="020F0502020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="E10002FF" w:usb1="4000ACFF" w:usb2="00000009" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Times New Roman">
    <w:panose1 w:val="02020603050405020304"/>
    <w:charset w:val="00"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="E0002AFF" w:usb1="C000785B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Arial">
    <w:panose1 w:val="020B0604020202020204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="E0002AFF" w:usb1="C000247B" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Courier New">
    <w:panose1 w:val="02070309020205020404"/>
    <w:charset w:val="00"/>
    <w:family w:val="modern"/>
    <w:pitch w:val="fixed"/>
    <w:sig w:usb0="E0002AFF" w:usb1="C0007843" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Calibri Light">
    <w:panose1 w:val="020F0302020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="E10002FF" w:usb1="4000ACFF" w:usb2="00000009" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/>
  </w:font>
  <w:font w:name="Georgia">
    <w:panose1 w:val="02040502050204030303"/>
    <w:charset w:val="00"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="E0002AFF" w:usb1="00000000" w:usb2="00000000" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
  </w:font>
</w:fonts>`;
  }

  /**
   * Generates word/settings.xml
   * Required for DOCX compliance - defines document settings
   */
  generateSettings(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
    <w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
    <w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
  </w:compat>
  <w:themeFontLang w:val="en-US"/>
  <w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>
</w:settings>`;
  }

  /**
   * Generates word/theme/theme1.xml
   * Required for DOCX compliance - defines color and font theme
   */
  generateTheme(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="4472C4"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri" panose="020F0502020204030204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>
            <a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs>
            <a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
              <a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>
            <a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
  <a:extLst>
    <a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">
      <thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>
    </a:ext>
  </a:extLst>
</a:theme>`;
  }
}
