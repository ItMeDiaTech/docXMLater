# DocXMLater API Methods Inventory

## Document Class

### Getters
- `getProperties()` - Get document properties
- `getParagraphs()` - Get all paragraphs
- `getTables()` - Get all tables
- `getTableOfContentsElements()` - Get TOC elements
- `getBodyElements()` - Get all body elements (paragraphs, tables, TOCs)
- `getParagraphCount()` - Count paragraphs
- `getTableCount()` - Count tables
- `getSection()` - Get section properties
- `getStyle(styleId)` - Get specific style
- `getStyles()` - Get all styles ✨ NEW
- `getStylesXml()` - Get raw styles XML ✨ NEW
- `getStylesManager()` - Get styles manager instance
- `getZipHandler()` - Get ZIP handler
- `getNumberingManager()` - Get numbering manager
- `getHeaderFooterManager()` - Get header/footer manager
- `getImageManager()` - Get image manager
- `getRelationshipManager()` - Get relationship manager
- `getBookmarkManager()` - Get bookmark manager
- `getBookmark(name)` - Get specific bookmark
- `getRevisionManager()` - Get revision manager
- `getRevisionStats()` - Get revision statistics
- `getCommentManager()` - Get comment manager
- `getComment(id)` - Get specific comment
- `getAllComments()` - Get all comments
- `getCommentStats()` - Get comment statistics
- `getCommentThread(commentId)` - Get comment with replies
- `getRecentComments(count)` - Get recent comments
- `getRecentRevisions(count)` - Get recent revisions
- `getAllInsertions()` - Get all insertion revisions
- `getAllDeletions()` - Get all deletion revisions
- `getParseWarnings()` - Get parsing warnings
- `getSizeStats()` - Get document size statistics
- `getPart(partName)` - Get document part
- `getContentTypes()` - Get content types map
- `getAllRelationships()` - Get all relationships
- `getAllRuns()` - Get all runs in the document
- `getFormattingReport()` - Get formatting statistics report

### Setters
- `setProperties(properties)` - Set document properties
- `setSection(section)` - Set section properties
- `setPageSize(width, height, orientation?)` - Set page size
- `setPageOrientation(orientation)` - Set page orientation
- `setMargins(margins)` - Set page margins
- `setHeader(header)` - Set default header
- `setFirstPageHeader(header)` - Set first page header
- `setEvenPageHeader(header)` - Set even page header
- `setFooter(footer)` - Set default footer
- `setFirstPageFooter(footer)` - Set first page footer
- `setEvenPageFooter(footer)` - Set even page footer
- `setStylesXml(xml)` - Set raw styles XML
- `setPart(partName, content)` - Set document part
- `setAuthor(author)` - Set document author (alias for setCreator)
- `setAllRunsFont(fontName)` - Apply font to all runs
- `setAllRunsSize(size)` - Apply size to all runs
- `setAllRunsColor(color)` - Apply color to all runs

### Other Methods
- `create()` - Static: Create new document
- `createEmpty()` - Static: Create minimal document ✨ NEW
- `load(filePath)` - Static: Load from file
- `loadFromBuffer(buffer)` - Static: Load from buffer
- `save(filePath)` - Save to file
- `toBuffer()` - Save to buffer
- `createParagraph(text?)` - Create paragraph
- `createTable(rows, cols)` - Create table
- `createHeader(type)` - Create header
- `createFooter(type)` - Create footer
- `createComment(author, content, initials?)` - Create comment
- `createReply(parentId, author, content, initials?)` - Create comment reply
- `createBookmark(name, startPara, endPara?)` - Create bookmark
- `addStyle(style)` - Add style
- `hasStyle(styleId)` - Check if style exists
- `removeStyle(styleId)` - Remove style ✨ NEW
- `updateStyle(styleId, properties)` - Update style ✨ NEW
- `addImage(source, width?, height?)` - Add image
- `addTable(table)` - Add table
- `addSection(section)` - Add section
- `addBulletList()` - Add bullet list
- `addNumberedList()` - Add numbered list
- `addMultiLevelList()` - Add multi-level list
- `addCommentToParagraph(para, comment/author, content?, initials?)` - Add comment
- `addBookmarkToParagraph(name, para)` - Add bookmark
- `addInsertionToParagraph(para, author, text, date?)` - Add insertion
- `addDeletionToParagraph(para, author, text, date?)` - Add deletion
- `updateHyperlinkUrls(urlMap)` - Update hyperlink URLs
- `clear()` - Clear document content
- `removePart(partName)` - Remove document part ✨ NEW
- `listParts()` - List all parts ✨ NEW
- `partExists(partName)` - Check if part exists ✨ NEW
- `addContentType(partNameOrExt, contentType)` - Add content type ✨ NEW
- `isTrackingChanges()` - Check if has revisions
- `updateStylesXml()` - Update styles XML in ZIP
- `findParagraphsByText(pattern)` - Find paragraphs by text pattern
- `getRunsByFont(fontName)` - Get runs using specific font
- `getRunsByColor(color)` - Get runs using specific color
- `getParagraphsByStyle(styleId)` - Get paragraphs by style

## Paragraph Class

### Getters
- `getText()` - Get paragraph text
- `getContent()` - Get content array (runs, hyperlinks, etc.)
- `getRuns()` - Get text runs
- `getAlignment()` - Get alignment
- `getFormatting()` - Get formatting properties
- `getNumbering()` - Get numbering info
- `getStyle()` - Get style ID

### Setters
- `setText(text)` - Set plain text
- `setAlignment(alignment)` - Set alignment
- `setIndentation(indentation)` - Set indentation
- `setSpacing(spacing)` - Set spacing
- `setSpaceBefore(twips)` - Set space before
- `setSpaceAfter(twips)` - Set space after
- `setLineSpacing(twips, rule?)` - Set line spacing
- `setStyle(styleId)` - Set style
- `setNumbering(numId, ilvl)` - Set numbering
- `setKeepNext(keep)` - Set keep with next
- `setKeepLines(keep)` - Set keep lines together
- `setPageBreakBefore(breakBefore)` - Set page break before

### Other Methods
- `addText(text, formatting?)` - Add text run
- `addRun(run)` - Add run
- `addHyperlink(hyperlink)` - Add hyperlink
- `addField(field)` - Add field
- `addRevision(revision)` - Add revision
- `addBookmarkStart(bookmark)` - Add bookmark start
- `addBookmarkEnd(bookmark)` - Add bookmark end
- `addCommentStart(comment)` - Add comment start
- `addCommentEnd(comment)` - Add comment end
- `clear()` - Clear content
- `isEmpty()` - Check if empty
- `isNumbered()` - Check if numbered
- `hasRevisions()` - Check if has revisions

## Run Class

### Getters
- `getText()` - Get text content
- `getFormatting()` - Get formatting properties

### Setters
- `setText(text)` - Set text content
- `setBold(bold)` - Set bold
- `setItalic(italic)` - Set italic
- `setUnderline(style)` - Set underline
- `setStrike(strike)` - Set strikethrough
- `setFont(font)` - Set font family
- `setSize(size)` - Set font size
- `setColor(color)` - Set text color
- `setHighlight(color)` - Set highlight color

### Other Methods
- `create(text, formatting?)` - Static: Create run
- `isRevision()` - Check if revision
- `isEmpty()` - Check if empty

## Style Class

### Getters
- `getStyleId()` - Get style ID
- `getName()` - Get style name
- `getType()` - Get style type
- `getBasedOn()` - Get parent style ID
- `getNext()` - Get next style ID
- `getProperties()` - Get all properties
- `getParagraphFormatting()` - Get paragraph formatting
- `getRunFormatting()` - Get run formatting

### Setters
- `setName(name)` - Set style name
- `setBasedOn(styleId)` - Set parent style
- `setNext(styleId)` - Set next style
- `setParagraphFormatting(formatting)` - Set paragraph formatting
- `setRunFormatting(formatting)` - Set run formatting

### Other Methods
- `create(properties)` - Static: Create style
- `createNormalStyle()` - Static: Create Normal style
- `createHeadingStyle(level)` - Static: Create heading style
- `createTitleStyle()` - Static: Create Title style
- `createSubtitleStyle()` - Static: Create Subtitle style
- `createListParagraphStyle()` - Static: Create list style
- `createTOCHeadingStyle()` - Static: Create TOC heading style
- `isDefault()` - Check if default style
- `isCustom()` - Check if custom style
- `isValid()` - Validate style
- `clone()` - Clone style with new ID
- `reset()` - Reset to minimal state

## Table Class

### Getters
- `getRow(index)` - Get specific row
- `getRows()` - Get all rows
- `getRowCount()` - Get row count
- `getCell(row, col)` - Get specific cell
- `getFormatting()` - Get formatting properties
- `getColumnCount()` - Get maximum column count

### Setters
- `setWidth(twips)` - Set table width
- `setAlignment(alignment)` - Set alignment
- `setLayout(layout)` - Set layout type
- `setBorders(borders)` - Set borders
- `setAllBorders(border)` - Set all borders
- `setCellSpacing(twips)` - Set cell spacing
- `setIndent(twips)` - Set indentation
- `setColumnWidths(widths)` - Set column widths

### Other Methods
- `addRow()` - Add new row
- `mergeCells(startRow, startCol, endRow, endCol)` - Merge cells
- `insertRow(index, row?)` - Insert row at position
- `removeRow(index)` - Remove row
- `sortRows(column, options?)` - Sort rows by column content
- `clone()` - Create deep copy

## Image Class

### Getters
- `getRelationshipId()` - Get relationship ID
- `getWidth()` - Get width in EMUs
- `getHeight()` - Get height in EMUs
- `getName()` - Get image name
- `getDescription()` - Get description

### Setters
- `setRelationshipId(id)` - Set relationship ID
- `setDimensions(width, height)` - Set dimensions
- `setName(name)` - Set image name
- `setDescription(desc)` - Set description

### Other Methods
- `create(properties)` - Static: Create image
- `calculateDimensions(buffer, maxWidth?, maxHeight?)` - Calculate dimensions

## Hyperlink Class

### Getters
- `getUrl()` - Get URL
- `getAnchor()` - Get anchor (internal link)
- `getText()` - Get display text
- `getTooltip()` - Get tooltip
- `getRelationshipId()` - Get relationship ID
- `getRun()` - Get text run
- `getFormatting()` - Get text formatting

### Setters
- `setText(text)` - Set display text
- `setTooltip(tooltip)` - Set tooltip
- `setRelationshipId(id)` - Set relationship ID
- `setUrl(url)` - Set URL ✨ IMPROVED
- `setFormatting(formatting)` - Set text formatting

### Other Methods
- `createExternal(url, text, formatting?)` - Static: Create external link
- `createInternal(anchor, text, formatting?)` - Static: Create internal link
- `createWebLink(url, text?, formatting?)` - Static: Create web link
- `createEmail(email, text?, formatting?)` - Static: Create email link
- `isExternal()` - Check if external link
- `isInternal()` - Check if internal link

## StylesManager Class

### Getters
- `getStyle(styleId)` - Get specific style
- `getAllStyles()` - Get all styles
- `getStylesByType(type)` - Get styles by type
- `getStyleCount()` - Get style count

### Other Methods
- `create()` - Static: Create with built-in styles
- `createEmpty()` - Static: Create empty manager
- `addStyle(style)` - Add style
- `hasStyle(styleId)` - Check if style exists
- `removeStyle(styleId)` - Remove style
- `clear()` - Clear all styles
- `generateStylesXml()` - Generate styles XML
- `validate(xml)` - Validate styles XML

## Section Class

### Getters
- `getProperties()` - Get section properties
- `getLineNumbering()` - Get line numbering configuration

### Setters
- `setPageSize(width, height, orientation?)` - Set page size
- `setOrientation(orientation)` - Set page orientation
- `setMargins(margins)` - Set page margins
- `setColumns(count, space?)` - Set column layout
- `setSectionType(type)` - Set section break type
- `setPageNumbering(start?, format?)` - Set page numbering
- `setTitlePage(titlePage?)` - Enable different first page
- `setHeaderReference(type, rId)` - Set header reference
- `setFooterReference(type, rId)` - Set footer reference
- `setVerticalAlignment(alignment)` - Set vertical alignment
- `setPaperSource(first?, other?)` - Set paper source
- `setColumnSeparator(separator?)` - Show column separator
- `setColumnWidths(widths)` - Set custom column widths
- `setTextDirection(direction)` - Set text direction
- `setLineNumbering(options)` - Set line numbering

### Other Methods
- `clearLineNumbering()` - Clear line numbering
- `clone()` - Create deep copy
- `create(properties?)` - Static: Create section
- `createLetter()` - Static: Create letter-sized section
- `createA4()` - Static: Create A4-sized section
- `createLandscape(pageSize?)` - Static: Create landscape section

## TableCell Class

### Getters
- `getParagraphs()` - Get all paragraphs
- `getText()` - Get combined text content
- `getFields()` - Get all fields
- `getFormatting()` - Get cell formatting

### Setters
- `setWidth(twips)` - Set cell width
- `setBorders(borders)` - Set cell borders
- `setShading(shading)` - Set background shading
- `setVerticalAlignment(alignment)` - Set vertical text alignment
- `setColumnSpan(span)` - Set horizontal merge span
- `setMargins(margins)` - Set cell margins
- `setAllMargins(margin)` - Set all margins uniformly
- `setTextDirection(direction)` - Set text direction
- `setFitText(fit?)` - Set fit text to width
- `setNoWrap(noWrap?)` - Set no text wrap
- `setVerticalMerge(merge)` - Set vertical merge
- `setTextAlignment(alignment)` - Set alignment for all paragraphs
- `setAllParagraphsStyle(styleId)` - Apply style to all paragraphs
- `setAllRunsFont(fontName)` - Apply font to all runs
- `setAllRunsSize(size)` - Apply size to all runs
- `setAllRunsColor(color)` - Apply color to all runs

### Other Methods
- `addParagraph(paragraph)` - Add paragraph
- `createParagraph(text?)` - Create and add paragraph
- `removeParagraph(index)` - Remove paragraph
- `addParagraphAt(index, paragraph)` - Insert paragraph at position
- `findFields(predicate)` - Find fields matching criteria
- `removeAllFields()` - Remove all fields
- `create(formatting?)` - Static: Create cell

## Comment Class

### Getters
- `getId()` - Get comment ID
- `getAuthor()` - Get author name
- `getInitials()` - Get author initials
- `getDate()` - Get comment date
- `getParentId()` - Get parent comment ID (for replies)
- `getRuns()` - Get text runs
- `getText()` - Get comment text
- `getContent()` - Get comment content (alias for getText)

### Setters
- `setId(id)` - Set comment ID (internal)
- `setAuthor(author)` - Set author name
- `setInitials(initials)` - Set author initials
- `setDate(date)` - Set comment date

### Other Methods
- `isReply()` - Check if this is a reply
- `isResolved()` - Check if comment is resolved/done
- `resolve()` - Mark comment as resolved
- `unresolve()` - Mark comment as unresolved
- `addRun(run)` - Add text run
- `create(author, content, initials?)` - Static: Create comment
- `createReply(parentId, author, content, initials?)` - Static: Create reply
- `createFormatted(author, runs, initials?)` - Static: Create with formatted runs

## CommentManager Class

### Getters
- `getComment(id)` - Get comment by ID
- `getAllComments()` - Get all top-level comments
- `getAllCommentsWithReplies()` - Get all comments including replies
- `getReplies(commentId)` - Get replies to a comment
- `getAuthors()` - Get all unique authors
- `getCommentsByAuthor(author)` - Get comments by author
- `getCommentsByDateRange(start, end)` - Get comments in date range
- `getCount()` - Get total comment count
- `getTopLevelCount()` - Get top-level count
- `getResolvedComments()` - Get resolved comments
- `getUnresolvedComments()` - Get unresolved comments
- `getRecentComments(count)` - Get most recent comments
- `getStats()` - Get comment statistics

### Other Methods
- `register(comment)` - Register comment with ID
- `registerExisting(comment)` - Register pre-existing comment
- `linkReplies()` - Link reply comments to parents
- `hasReplies(commentId)` - Check if has replies
- `removeComment(id)` - Remove comment and replies
- `clear()` - Clear all comments
- `createComment(author, content, initials?)` - Create and register comment
- `createReply(parentId, author, content, initials?)` - Create and register reply
- `isEmpty()` - Check if empty
- `getCommentThread(commentId)` - Get comment with all replies
- `findCommentsByText(searchText)` - Search by content
- `generateCommentsXml()` - Generate comments.xml
- `create()` - Static: Create manager

## Methods Status Review (Updated December 2025)

### Document Class - Implemented Methods
- [x] `getWordCount()` - Count words in document ✓
- [x] `getCharacterCount()` - Count characters ✓
- [x] `getImages()` - Get all images in document ✓
- [x] `findText(text)` - Search for text ✓
- [x] `replaceText(find, replace)` - Replace text globally ✓
- [x] `getHyperlinks()` - Get all hyperlinks ✓
- [x] `getBookmarks()` - Get all bookmarks ✓
- [x] `getFields()` - Get all fields ✓ NEW

### Paragraph Class - Implemented Methods
- [x] `clone()` - Clone paragraph ✓
- [x] `setBorder(border)` - Set paragraph border ✓
- [x] `setShading(shading)` - Set paragraph shading ✓

### Table Class - Implemented Methods
- [x] `removeRow(index)` - Remove row ✓
- [x] `insertRow(index)` - Insert row at position ✓
- [x] `getColumnCount()` - Get column count ✓

### Run Class - Implemented Methods
- [x] `setSubscript(sub)` - Set subscript ✓
- [x] `setSuperscript(super)` - Set superscript ✓
- [x] `setSmallCaps(small)` - Set small caps ✓
- [x] `setAllCaps(all)` - Set all caps ✓

---

## Potential Future Enhancements

### Document Class
- [ ] `getPageCount()` - Calculate page count (complex - requires layout engine)
- [ ] `setLanguage(lang)` - Set document language

### Table Class
- [ ] `addColumn()` - Add column (complex - requires synchronized grid updates)
- [ ] `removeColumn(index)` - Remove column (very complex - cannot remove if cells span)

### Image Class
- [ ] `rotate(degrees)` - Rotate image
- [ ] `crop(rect)` - Crop image

### Bookmark Class
- [ ] `getRange()` - Get text range
- [ ] `moveTo(para)` - Move bookmark

### Header/Footer Class
- [ ] `getDifferentFirstPage()` - Check first page setting
- [ ] `getDifferentOddEven()` - Check odd/even setting
- [ ] `getLinkToPrevious()` - Check link to previous

### NumberingManager Class
- [ ] `getNumberingStyles()` - Get all numbering styles
- [ ] `resetNumbering()` - Reset numbering
- [ ] `continueNumbering()` - Continue from previous

## Summary

The API provides comprehensive coverage for document creation and editing tasks:

- Core document manipulation (create, load, save, properties)
- Text and paragraph formatting (bold, italic, colors, fonts, alignment, spacing)
- Advanced formatting (borders, shading, styles)
- Search and replace with regex support
- Formatting analysis and bulk operations
- Word/character counting
- Styles management (create, clone, reset, apply)
- Tables (create, modify, merge, sort, column widths)
- Sections (page setup, margins, line numbering, columns)
- Headers and footers (different first/odd/even pages)
- Comments with resolution tracking
- Track changes (revisions, acceptance)
- Bookmarks and cross-references
- Images (embed, position, wrap)
- Hyperlinks (internal/external)
- Fields (merge, date, page numbers, TOC)
- Low-level document parts access

Remaining gaps are primarily in:
- Advanced image manipulation (rotation, cropping)
- Table column add/remove operations
- Page counting (requires layout engine)
- Header/footer introspection methods