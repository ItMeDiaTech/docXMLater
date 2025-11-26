import { ZipHandler } from '../zip/ZipHandler';

/**
 * Accepts all tracked changes in a Word document per Microsoft's OpenXML SDK pattern
 * 
 * This implementation follows the official Microsoft approach:
 * 1. Insertions (<w:ins>): Keep content, remove wrapper tags
 * 2. Deletions (<w:del>): Remove entirely (content and tags)
 * 3. Move From (<w:moveFrom>): Remove entirely (source of move)
 * 4. Move To (<w:moveTo>): Keep content, remove wrapper (destination of move)
 * 5. Property changes: Remove all *Change elements
 * 6. Range markers: Remove all boundary markers
 * 
 * Also cleans up metadata in people.xml, settings.xml, and core.xml
 * 
 * @see https://learn.microsoft.com/en-us/office/open-xml/how-to-accept-all-revisions
 */
export class RevisionAcceptor {
  private zipHandler: ZipHandler;

  constructor(zipHandler: ZipHandler) {
    this.zipHandler = zipHandler;
  }

  /**
   * Main method to accept all revisions in the document
   */
  public async acceptAllRevisions(): Promise<void> {
    // Process document.xml
    await this.processDocumentPart('word/document.xml');
    
    // Process headers
    const files = this.zipHandler.getFilePaths();
    for (const file of files) {
      if (file.match(/^word\/header\d+\.xml$/)) {
        await this.processDocumentPart(file);
      }
      if (file.match(/^word\/footer\d+\.xml$/)) {
        await this.processDocumentPart(file);
      }
    }

    // Clean up metadata files
    this.cleanupPeopleXml();
    this.cleanupSettingsXml();
    this.cleanupCorePropsXml();
  }

  /**
   * Process a document part (document.xml, header, footer) to accept revisions
   */
  private async processDocumentPart(partPath: string): Promise<void> {
    const xml = this.zipHandler.getFileAsString(partPath);
    if (!xml) {
      return;
    }

    let content = xml;

    // Step 1: Remove all range markers FIRST (before processing revisions)
    // This prevents orphaned references when revision content is modified
    content = this.removeAllRangeMarkers(content);

    // Step 2: Remove all property change elements
    // These track formatting changes and must be removed before other processing
    content = this.removeAllPropertyChanges(content);

    // Step 3: Process deletions - remove entire element INCLUDING content
    // Must be done before insertions to handle nested scenarios
    content = this.acceptDeletions(content);

    // Step 4: Process move operations
    // Remove moveFrom entirely (source), unwrap moveTo (destination)
    content = this.acceptMoveFrom(content);
    content = this.acceptMoveTo(content);

    // Step 5: Process insertions - keep content, remove wrapper
    content = this.acceptInsertions(content);

    // Step 6: Final cleanup - remove any remaining orphaned tags
    content = this.cleanupOrphanedTags(content);

    // Update the file
    this.zipHandler.updateFile(partPath, content);
  }

  /**
   * Remove all range marker elements
   * These are boundary markers for tracked changes and moves
   */
  private removeAllRangeMarkers(xml: string): string {
    const patterns = [
      // Move range markers
      /<w:moveFromRangeStart[^>]*(?:\/>|>.*?<\/w:moveFromRangeStart>)/gs,
      /<w:moveFromRangeEnd[^>]*(?:\/>|>.*?<\/w:moveFromRangeEnd>)/gs,
      /<w:moveToRangeStart[^>]*(?:\/>|>.*?<\/w:moveToRangeStart>)/gs,
      /<w:moveToRangeEnd[^>]*(?:\/>|>.*?<\/w:moveToRangeEnd>)/gs,
      // Custom XML range markers
      /<w:customXmlInsRangeStart[^>]*(?:\/>|>.*?<\/w:customXmlInsRangeStart>)/gs,
      /<w:customXmlInsRangeEnd[^>]*(?:\/>|>.*?<\/w:customXmlInsRangeEnd>)/gs,
      /<w:customXmlDelRangeStart[^>]*(?:\/>|>.*?<\/w:customXmlDelRangeStart>)/gs,
      /<w:customXmlDelRangeEnd[^>]*(?:\/>|>.*?<\/w:customXmlDelRangeEnd>)/gs,
      /<w:customXmlMoveFromRangeStart[^>]*(?:\/>|>.*?<\/w:customXmlMoveFromRangeStart>)/gs,
      /<w:customXmlMoveFromRangeEnd[^>]*(?:\/>|>.*?<\/w:customXmlMoveFromRangeEnd>)/gs,
      /<w:customXmlMoveToRangeStart[^>]*(?:\/>|>.*?<\/w:customXmlMoveToRangeStart>)/gs,
      /<w:customXmlMoveToRangeEnd[^>]*(?:\/>|>.*?<\/w:customXmlMoveToRangeEnd>)/gs,
    ];

    let result = xml;
    for (const pattern of patterns) {
      result = result.replace(pattern, '');
    }
    return result;
  }

  /**
   * Remove all property change tracking elements
   * Per ECMA-376, these track previous state of formatting
   */
  private removeAllPropertyChanges(xml: string): string {
    const patterns = [
      // Run property changes
      /<w:rPrChange[^>]*>[\s\S]*?<\/w:rPrChange>/g,
      // Paragraph property changes  
      /<w:pPrChange[^>]*>[\s\S]*?<\/w:pPrChange>/g,
      // Table property changes
      /<w:tblPrChange[^>]*>[\s\S]*?<\/w:tblPrChange>/g,
      /<w:tblPrExChange[^>]*>[\s\S]*?<\/w:tblPrExChange>/g,
      // Table cell property changes
      /<w:tcPrChange[^>]*>[\s\S]*?<\/w:tcPrChange>/g,
      // Table row property changes
      /<w:trPrChange[^>]*>[\s\S]*?<\/w:trPrChange>/g,
      // Section property changes
      /<w:sectPrChange[^>]*>[\s\S]*?<\/w:sectPrChange>/g,
      // Table grid changes
      /<w:tblGridChange[^>]*>[\s\S]*?<\/w:tblGridChange>/g,
      // Numbering changes
      /<w:numberingChange[^>]*>[\s\S]*?<\/w:numberingChange>/g,
    ];

    let result = xml;
    for (const pattern of patterns) {
      result = result.replace(pattern, '');
    }
    return result;
  }

  /**
   * Accept deletions - remove the entire <w:del> element including its content
   * 
   * Per Microsoft SDK: "DeletedRun elements should be removed along with their content"
   */
  private acceptDeletions(xml: string): string {
    let result = xml;
    let previousLength = 0;
    
    // Iterate until no more deletions (handles nested cases)
    while (result.length !== previousLength) {
      previousLength = result.length;
      
      // Match complete <w:del ...>...</w:del> elements and remove entirely
      result = result.replace(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/g, '');
    }
    
    // Also remove self-closing deletion tags
    result = result.replace(/<w:del\b[^>]*\/>/g, '');
    
    return result;
  }

  /**
   * Accept moveFrom - remove the entire element (source of moved content)
   * 
   * The content exists at the moveTo destination, so we discard the source
   */
  private acceptMoveFrom(xml: string): string {
    let result = xml;
    let previousLength = 0;
    
    while (result.length !== previousLength) {
      previousLength = result.length;
      result = result.replace(/<w:moveFrom\b[^>]*>[\s\S]*?<\/w:moveFrom>/g, '');
    }
    
    // Also remove self-closing tags
    result = result.replace(/<w:moveFrom\b[^>]*\/>/g, '');
    
    return result;
  }

  /**
   * Accept moveTo - keep the content, remove the wrapper tags
   * 
   * The moveTo location is where the content should remain
   */
  private acceptMoveTo(xml: string): string {
    let result = xml;
    
    // Remove closing tags first (prevents issues with regex matching)
    result = result.replace(/<\/w:moveTo>/g, '');
    
    // Remove opening tags (keeps content that was inside)
    result = result.replace(/<w:moveTo\b[^>]*>/g, '');
    
    return result;
  }

  /**
   * Accept insertions - keep the content, remove the wrapper tags
   *
   * Per Microsoft SDK: "InsertedRun elements should be unwrapped, keeping their content"
   *
   * IMPORTANT: This method now handles relationship ID remapping for images inside insertions.
   * When Word tracks changes with images, it can reuse relationship IDs (like rId5) because
   * they're in separate tracked change contexts. But when we unwrap them, duplicate IDs
   * cause corruption. This method assigns new unique IDs to images inside insertions.
   */
  private acceptInsertions(xml: string): string {
    let result = xml;

    // Parse existing relationships
    const relationships = this.parseRelationships();
    const existingIds = new Set(relationships.keys());

    // Process each w:ins element and remap images one by one
    const insRegex = /<w:ins\b[^>]*>[\s\S]*?<\/w:ins>/g;

    result = result.replace(insRegex, (insMatch) => {
      // For each image reference inside this insertion, generate a unique new ID
      return insMatch.replace(/r:embed="(rId\d+)"/g, (embedMatch, oldId) => {
        // Generate new unique ID for THIS occurrence
        const newId = this.getNextRelationshipId(existingIds);
        existingIds.add(newId);

        // Add relationship with same target as original
        const target = relationships.get(oldId);
        if (target) {
          this.addRelationship(
            newId,
            target,
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
          );
        }

        return `r:embed="${newId}"`;
      });
    });

    // Now unwrap the w:ins tags (content has unique remapped IDs)
    result = result.replace(/<\/w:ins>/g, '');
    result = result.replace(/<w:ins\b[^>]*>/g, '');

    return result;
  }

  /**
   * Final cleanup to remove any orphaned or malformed revision-related tags
   */
  private cleanupOrphanedTags(xml: string): string {
    let result = xml;

    // Remove any remaining self-closing revision tags
    result = result.replace(/<w:ins\b[^>]*\/>/g, '');
    result = result.replace(/<w:del\b[^>]*\/>/g, '');
    result = result.replace(/<w:moveFrom\b[^>]*\/>/g, '');
    result = result.replace(/<w:moveTo\b[^>]*\/>/g, '');

    // Remove empty w:r elements that might be left after removing deletions
    result = result.replace(/<w:r\b[^>]*>\s*<\/w:r>/g, '');

    // Remove empty w:p elements (but preserve those with properties)
    // We keep <w:p><w:pPr>...</w:pPr></w:p> as those are intentional empty paragraphs with styling
    result = result.replace(/<w:p>\s*<\/w:p>/g, '');

    return result;
  }

  /**
   * Parse relationship IDs from word/_rels/document.xml.rels
   * Returns a map of relationship ID to target path
   */
  private parseRelationships(): Map<string, string> {
    const relsXml = this.zipHandler.getFileAsString('word/_rels/document.xml.rels');
    if (!relsXml) return new Map();

    const map = new Map<string, string>();
    const relationshipRegex = /<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/>/g;
    let match;

    while ((match = relationshipRegex.exec(relsXml)) !== null) {
      if (match[1] && match[2]) {
        map.set(match[1], match[2]); // rId -> target path
      }
    }

    return map;
  }

  /**
   * Get the next available relationship ID
   * Finds the highest numeric ID and increments it
   */
  private getNextRelationshipId(existingIds: Set<string>): string {
    let maxId = 0;
    for (const id of existingIds) {
      const num = parseInt(id.replace('rId', ''));
      if (!isNaN(num) && num > maxId) {
        maxId = num;
      }
    }
    return `rId${maxId + 1}`;
  }

  /**
   * Add a new relationship to word/_rels/document.xml.rels
   */
  private addRelationship(rId: string, target: string, type: string): void {
    const relsXml = this.zipHandler.getFileAsString('word/_rels/document.xml.rels');
    if (!relsXml) return;

    // Insert new relationship before closing tag
    const newRel = `<Relationship Id="${rId}" Type="${type}" Target="${target}"/>`;
    const updated = relsXml.replace('</Relationships>', `${newRel}\n</Relationships>`);

    this.zipHandler.updateFile('word/_rels/document.xml.rels', updated);
  }

  /**
   * Clean up word/people.xml - remove all revision authors
   * 
   * Handles both w: and w15: namespace variants
   */
  private cleanupPeopleXml(): void {
    const peopleXml = this.zipHandler.getFileAsString('word/people.xml');
    if (!peopleXml) {
      return;
    }

    let content = peopleXml;

    // Remove all person elements in any namespace variant
    content = content.replace(/<w:person\b[^>]*>[\s\S]*?<\/w:person>/g, '');
    content = content.replace(/<w15:person\b[^>]*>[\s\S]*?<\/w15:person>/g, '');
    
    // Handle any namespace-prefixed variants (w1:, w2:, etc.)
    content = content.replace(/<w\d+:person\b[^>]*>[\s\S]*?<\/w\d+:person>/g, '');
    
    // Also remove self-closing person elements
    content = content.replace(/<w:person\b[^>]*\/>/g, '');
    content = content.replace(/<w15:person\b[^>]*\/>/g, '');
    content = content.replace(/<w\d+:person\b[^>]*\/>/g, '');

    this.zipHandler.updateFile('word/people.xml', content);
  }

  /**
   * Clean up word/settings.xml - disable track changes
   */
  private cleanupSettingsXml(): void {
    const settingsXml = this.zipHandler.getFileAsString('word/settings.xml');
    if (!settingsXml) {
      return;
    }

    let content = settingsXml;

    // Remove trackRevisions element (enables tracking)
    content = content.replace(/<w:trackRevisions\b[^>]*\/>/g, '');
    content = content.replace(/<w:trackRevisions\b[^>]*>[\s\S]*?<\/w:trackRevisions>/g, '');

    // Remove revisionView element (controls which revisions are visible)
    content = content.replace(/<w:revisionView\b[^>]*\/>/g, '');
    content = content.replace(/<w:revisionView\b[^>]*>[\s\S]*?<\/w:revisionView>/g, '');

    // Remove doNotTrackMoves (prevents move tracking)
    content = content.replace(/<w:doNotTrackMoves\b[^>]*\/>/g, '');
    content = content.replace(/<w:doNotTrackMoves\b[^>]*>[\s\S]*?<\/w:doNotTrackMoves>/g, '');

    // Remove doNotTrackFormatting
    content = content.replace(/<w:doNotTrackFormatting\b[^>]*\/>/g, '');
    content = content.replace(/<w:doNotTrackFormatting\b[^>]*>[\s\S]*?<\/w:doNotTrackFormatting>/g, '');

    this.zipHandler.updateFile('word/settings.xml', content);
  }

  /**
   * Clean up docProps/core.xml - reset revision count
   */
  private cleanupCorePropsXml(): void {
    const coreXml = this.zipHandler.getFileAsString('docProps/core.xml');
    if (!coreXml) {
      return;
    }

    // Reset revision count to 1
    const content = coreXml.replace(
      /<cp:revision>\d+<\/cp:revision>/g, 
      '<cp:revision>1</cp:revision>'
    );

    this.zipHandler.updateFile('docProps/core.xml', content);
  }
}

/**
 * Convenience function to accept all revisions in a document
 */
export async function acceptAllRevisions(zipHandler: ZipHandler): Promise<void> {
  const acceptor = new RevisionAcceptor(zipHandler);
  await acceptor.acceptAllRevisions();
}
