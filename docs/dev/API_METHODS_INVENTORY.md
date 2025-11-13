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
- `getPart(partName)` - Get document part ✨ NEW
- `getContentTypes()` - Get content types map ✨ NEW
- `getAllRelationships()` - Get all relationships ✨ NEW

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
- `setStylesXml(xml)` - Set raw styles XML ✨ NEW
- `setPart(partName, content)` - Set document part ✨ NEW

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
- `isValid()` - Validate style ✨ NEW

## Table Class

### Getters
- `getRow(index)` - Get specific row
- `getRows()` - Get all rows
- `getRowCount()` - Get row count
- `getCell(row, col)` - Get specific cell
- `getFormatting()` - Get formatting properties

### Setters
- `setWidth(twips)` - Set table width
- `setAlignment(alignment)` - Set alignment
- `setLayout(layout)` - Set layout type
- `setBorders(borders)` - Set borders
- `setAllBorders(border)` - Set all borders
- `setCellSpacing(twips)` - Set cell spacing
- `setIndent(twips)` - Set indentation

### Other Methods
- `addRow()` - Add new row
- `mergeCells(startRow, startCol, endRow, endCol)` - Merge cells

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
- `validate(xml)` - Validate styles XML ✨ NEW

## Potential Missing Methods to Consider

### Document Class
- [ ] `getPageCount()` - Calculate page count (complex)
- [ ] `getWordCount()` - Count words in document
- [ ] `getCharacterCount()` - Count characters
- [ ] `setLanguage(lang)` - Set document language
- [ ] `setAuthor(author)` - Convenience for setting creator
- [ ] `getImages()` - Get all images in document
- [ ] `findText(text)` - Search for text
- [ ] `replaceText(find, replace)` - Replace text globally
- [ ] `insertParagraphAt(index, para)` - Insert at specific position
- [ ] `removeParagraph(para/index)` - Remove specific paragraph
- [ ] `removeTable(table/index)` - Remove specific table
- [ ] `getHyperlinks()` - Get all hyperlinks
- [ ] `getBookmarks()` - Get all bookmarks
- [ ] `getFields()` - Get all fields
- [ ] `protect(password?)` - Protect document
- [ ] `unprotect(password?)` - Unprotect document
- [ ] `isProtected()` - Check protection status

### Paragraph Class
- [ ] `getLength()` - Get text length
- [ ] `getWordCount()` - Count words
- [ ] `setBorder(border)` - Set paragraph border
- [ ] `setShading(shading)` - Set paragraph shading
- [ ] `setTabs(tabs)` - Set tab stops
- [ ] `clone()` - Clone paragraph

### Table Class
- [ ] `removeRow(index)` - Remove row
- [ ] `insertRow(index)` - Insert row at position
- [ ] `addColumn()` - Add column
- [ ] `removeColumn(index)` - Remove column
- [ ] `getColumnCount()` - Get column count
- [ ] `setColumnWidths(widths)` - Set column widths
- [ ] `sortRows(column, ascending?)` - Sort table rows

### Style Class
- [ ] `clone()` - Clone style
- [ ] `inherit(parentStyle)` - Inherit from parent
- [ ] `reset()` - Reset to defaults

### Run Class
- [ ] `setSubscript(sub)` - Set subscript
- [ ] `setSuperscript(super)` - Set superscript
- [ ] `setSmallCaps(small)` - Set small caps
- [ ] `setAllCaps(all)` - Set all caps
- [ ] `setEmboss(emboss)` - Set emboss effect
- [ ] `setImprint(imprint)` - Set imprint effect
- [ ] `setShadow(shadow)` - Set shadow effect

### Section Class
- [ ] `getPageCount()` - Get pages in section
- [ ] `setLineNumbers(show)` - Set line numbering
- [ ] `setVerticalAlignment(align)` - Set vertical alignment

### Comment Class
- [ ] `resolve()` - Mark as resolved
- [ ] `isResolved()` - Check if resolved
- [ ] `setAuthorInitials(initials)` - Set initials

### Image Class
- [ ] `rotate(degrees)` - Rotate image
- [ ] `crop(rect)` - Crop image
- [ ] `setAltText(text)` - Set alt text
- [ ] `setWrapType(type)` - Set text wrapping

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

✅ **Strong Coverage:**
- Core document manipulation
- Text and paragraph formatting
- Styles management
- Tables
- Headers/footers
- Comments and revisions
- Bookmarks
- Images
- Hyperlinks
- Low-level document parts access

⚠️ **Potential Gaps:**
- Text search and replace
- Document protection
- Advanced image manipulation
- Paragraph borders and shading
- Tab stops
- Column operations in tables
- Page counting
- Word/character counting
- Resolved comments tracking
- Text wrapping for images

The API is quite comprehensive for most document creation and editing tasks. The main gaps are in advanced features like document protection, search/replace, and some formatting options that are less commonly used.