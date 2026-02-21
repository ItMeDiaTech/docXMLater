/**
 * Complete catalog of legacy boolean compat elements defined in
 * ECMA-376 4th Edition Part 4 (Transitional) and [MS-DOCX] specification.
 *
 * These elements are direct children of <w:compat> with complex type CT_OnOff.
 * They emulate rendering quirks of older word processors. When compatibilityMode
 * is upgraded to 15, Word's modern layout engine no longer needs them and they
 * should be removed.
 *
 * Source: ECMA-376 4th Edition Part 4 sections 2.15.3.1 through 2.15.3.65
 * and the [MS-DOCX] specification section 2.15.3.
 * Verified against: http://www.datypic.com/sc/ooxml/t-w_CT_Compat.html (65 elements)
 *
 * Element names include the w: namespace prefix as they appear in XML.
 */
export const LEGACY_COMPAT_ELEMENTS: string[] = [
  // Word 6/95 emulation
  'w:useSingleBorderforContiguousCells',
  'w:wpJustification',
  'w:noTabHangInd',
  'w:noLeading',
  'w:spaceForUL',
  'w:noColumnBalance',
  'w:balanceSingleByteDoubleByteWidth',
  'w:noExtraLineSpacing',
  'w:doNotLeaveBackslashAlone',
  'w:ulTrailSpace',
  'w:doNotExpandShiftReturn',
  'w:spacingInWholePoints',
  'w:lineWrapLikeWord6',
  'w:printBodyTextBeforeHeader',
  'w:printColBlack',
  'w:wpSpaceWidth',
  'w:showBreaksInFrames',
  'w:subFontBySize',
  'w:suppressBottomSpacing',
  'w:suppressTopSpacing',
  'w:suppressSpacingAtTopOfPage',
  'w:suppressTopSpacingWP',
  'w:suppressSpBfAfterPgBrk',
  'w:swapBordersFacingPages',
  'w:convMailMergeEsc',
  'w:truncateFontHeightsLikeWP6',
  'w:mwSmallCaps',

  // Word 97/2000 emulation
  'w:usePrinterMetrics',
  'w:doNotSuppressParagraphBorders',
  'w:wrapTrailSpaces',
  'w:footnoteLayoutLikeWW8',
  'w:shapeLayoutLikeWW8',
  'w:alignTablesRowByRow',
  'w:forgetLastTabAlignment',
  'w:adjustLineHeightInTable',
  'w:autoSpaceLikeWord95',
  'w:noSpaceRaiseLower',
  'w:doNotUseHTMLParagraphAutoSpacing',
  'w:layoutRawTableWidth',
  'w:layoutTableRowsApart',
  'w:useWord97LineBreakRules',
  'w:doNotBreakWrappedTables',
  'w:doNotSnapToGridInCell',
  'w:selectFldWithFirstOrLastChar',
  'w:applyBreakingRules',
  'w:doNotWrapTextWithPunct',
  'w:doNotUseEastAsianBreakRules',
  'w:useWord2002TableStyleRules',

  // Word 2003+ additions
  'w:growAutofit',
  'w:useFELayout',
  'w:useNormalStyleForList',
  'w:doNotUseIndentAsNumberingTabStop',
  'w:useAltKinsokuLineBreakRules',
  'w:allowSpaceOfSameStyleInTable',
  'w:doNotSuppressIndentation',
  'w:doNotAutofitConstrainedTables',
  'w:autofitToFirstFixedWidthCell',
  'w:underlineTabInNumList',
  'w:displayHangulFixedWidth',
  'w:splitPgBreakAndParaMark',
  'w:doNotVertAlignCellWithSp',
  'w:doNotBreakConstrainedForcedTable',
  'w:doNotVertAlignInTxbx',
  'w:useAnsiKerningPairs',
  'w:cachedColBalance',
];

/**
 * Set of legacy compat element names without the w: prefix,
 * for efficient lookup during parsing and filtering.
 */
export const LEGACY_COMPAT_ELEMENT_NAMES = new Set<string>(
  LEGACY_COMPAT_ELEMENTS.map(e => e.replace('w:', ''))
);

import type { CompatSetting } from '../types/compatibility-types';

/**
 * The w:compatSetting entries that Word 2013+ includes for
 * fully modern documents. These are the settings emitted when
 * a user clicks File > Info > Convert in Word.
 *
 * Note: useWord2013TrackBottomHyphenation is intentionally omitted.
 * Its default value varies across Word versions and its presence
 * could change hyphenation behavior. Real Word 2016+ documents
 * sometimes include it as "0" and sometimes as "1".
 */
export const MODERN_COMPAT_SETTINGS: CompatSetting[] = [
  {
    name: 'compatibilityMode',
    uri: 'http://schemas.microsoft.com/office/word',
    val: '15',
  },
  {
    name: 'overrideTableStyleFontSizeAndJustification',
    uri: 'http://schemas.microsoft.com/office/word',
    val: '1',
  },
  {
    name: 'enableOpenTypeFeatures',
    uri: 'http://schemas.microsoft.com/office/word',
    val: '1',
  },
  {
    name: 'doNotFlipMirrorIndents',
    uri: 'http://schemas.microsoft.com/office/word',
    val: '1',
  },
  {
    name: 'differentiateMultirowTableHeaders',
    uri: 'http://schemas.microsoft.com/office/word',
    val: '1',
  },
];

/** The Microsoft Office Word URI used for w:compatSetting entries */
export const MS_WORD_COMPAT_URI = 'http://schemas.microsoft.com/office/word';
