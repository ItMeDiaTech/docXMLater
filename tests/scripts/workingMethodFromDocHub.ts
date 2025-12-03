try {
      // ============================================
      // STEP 1: GET ALL EXISTING HEADINGS
      // ============================================
      interface HeadingInfo {
        paragraph: Paragraph;
        level: number;
        text: string;
        bookmark: Bookmark;
      }

      const allHeadings: HeadingInfo[] = [];
      const allParagraphs = doc.getAllParagraphs(); // Searches body AND tables

      this.log.debug(`Scanning ${allParagraphs.length} paragraphs for headings...`);

      for (const para of allParagraphs) {
        const style = para.getStyle();

        // FIX Issue 1: Match ANY heading level (not just 1-3)
        // This allows TOC field instruction filtering to work correctly
        // Previous regex /^Heading\s*([1-3])$/i only matched levels 1-3
        // New regex /^Heading\s*(\d+)$/i matches any numeric heading level
        const match = style?.match(/^Heading\s*(\d+)$/i);
        if (match && match[1]) {
          const level = parseInt(match[1], 10);
          const text = para.getText().trim();

          if (text) {
            // Create unique bookmark for this heading
            const bookmarkName = `_Heading_${level}_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
            const bookmark = new Bookmark({ name: bookmarkName });
            const registered = doc.getBookmarkManager().register(bookmark);

            // Add bookmark to the heading paragraph
            para.addBookmark(registered);

            allHeadings.push({
              paragraph: para,
              level: level,
              text: text,
              bookmark: registered,
            });

            this.log.debug(`Found Heading${level}: "${text}" (bookmark: ${bookmarkName})`);
          }
        }
      }

      this.log.info(`Found ${allHeadings.length} headings in document`);

      if (allHeadings.length === 0) {
        this.log.warn('No headings found - TOC cannot be populated');
        return 0;
      }

      // ============================================
      // STEP 2: GET ALL TOC ELEMENTS IN DOCUMENT
      // ============================================
      const tocElements = doc.getTableOfContentsElements();
      this.log.info(`Found ${tocElements.length} TOC element(s) in document`);

      if (tocElements.length === 0) {
        this.log.warn('No TOC elements found in document');
        return 0;
      }

      // ============================================
      // STEP 3: PROCESS EACH TOC
      // ============================================
      for (const tocElement of tocElements) {
        const toc = (tocElement as any).toc as TableOfContents;
        if (!toc) {
          this.log.warn('TOC element missing toc property, skipping');
          continue;
        }

        // Parse field instruction to determine which levels to include
        const fieldInstruction = toc.getFieldInstruction();
        this.log.debug(`TOC field instruction: ${fieldInstruction}`);

        // Extract which heading levels this TOC should include
        const levelsToInclude = this.parseTOCLevels(fieldInstruction);
        this.log.debug(`TOC includes heading levels: ${levelsToInclude.join(', ')}`);

        // Filter headings for this TOC
        const tocHeadings = allHeadings.filter((h) => levelsToInclude.includes(h.level));

        if (tocHeadings.length === 0) {
          this.log.warn('No headings match TOC level filter');
          continue;
        }

        this.log.info(
          `Building TOC with ${tocHeadings.length} headings (levels: ${levelsToInclude.join(', ')})`
        );

        // Calculate minimum level for relative indentation
        const minLevel = Math.min(...tocHeadings.map((h) => h.level));

        // ============================================
        // STEP 4: BUILD MANUAL TOC ENTRIES
        // ============================================
        const tocParagraphs: Paragraph[] = [];

        for (const heading of tocHeadings) {
          // Create paragraph for TOC entry
          const tocEntry = new Paragraph();

          // Set spacing: 0 before and 0 after
          tocEntry.setSpaceBefore(0);
          tocEntry.setSpaceAfter(0);

          // Set line spacing to 240 (single spacing)
          tocEntry.setLineSpacing(240, 'auto');

          // Set left alignment
          tocEntry.setAlignment('left');

          // Calculate indentation based on heading level (0.25" per level above minimum)
          // Level 2 when minLevel=2: 0", Level 3 when minLevel=2: 0.25" (360 twips)
          const indentTwips = (heading.level - minLevel) * 360;
          if (indentTwips > 0) {
            tocEntry.setLeftIndent(indentTwips);
          }

          // Create internal hyperlink
          const hyperlink = Hyperlink.createInternal(heading.bookmark.getName(), heading.text, {
            font: 'Verdana',
            size: 12,
            color: '0000FF', // Blue
            underline: 'single',
          });

          // Add hyperlink to paragraph
          tocEntry.addHyperlink(hyperlink);

          tocParagraphs.push(tocEntry);

          this.log.debug(
            `Created TOC entry for ${heading.text} (Level ${heading.level}, indent: ${indentTwips} twips)`
          );
        }

        // ============================================
        // STEP 5: INSERT TOC ENTRIES INTO DOCUMENT
        // ============================================
        const bodyElements = doc.getBodyElements();
        const tocIndex = bodyElements.indexOf(tocElement);

        if (tocIndex !== -1) {
          // Remove the TOC element
          doc.removeTocAt(tocIndex);
          this.log.debug(`Removed original TOC field at index ${tocIndex}`);

          // Insert all TOC entry paragraphs at that position
          for (let i = 0; i < tocParagraphs.length; i++) {
            doc.insertParagraphAt(tocIndex + i, tocParagraphs[i]!);
          }

          totalEntriesCreated += tocParagraphs.length;
          this.log.info(
            `Inserted ${tocParagraphs.length} TOC entries at index ${tocIndex} replacing TOC field`
          );
        } else {
          this.log.warn('Could not find TOC element in body, skipping TOC replacement');
        }
      }

      this.log.info(
        `Successfully created ${totalEntriesCreated} TOC entries with internal hyperlinks`
      );
      return totalEntriesCreated;
    } catch (error) {
      this.log.error(
        `Error in manual TOC population: ${error instanceof Error ? error.message : 'Unknown error'}`
      );
      // Log stack trace for debugging
      if (error instanceof Error && error.stack) {
        this.log.debug(`Stack trace: ${error.stack}`);
      }
      // Don't throw - allow document processing to continue
      return totalEntriesCreated;
    }
  }