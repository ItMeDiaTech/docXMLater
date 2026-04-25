/**
 * Shared type definitions for document settings (settings.xml)
 *
 * These types are used by both Document.ts (in-memory state) and
 * DocumentGenerator.ts (XML generation) to ensure consistency.
 */

/**
 * Document protection settings per ECMA-376 CT_DocProtect
 */
export interface DocumentProtection {
  edit: 'readOnly' | 'comments' | 'trackedChanges' | 'forms';
  enforcement: boolean;
  /**
   * `w:formatting` per CT_DocProtect §17.15.1.29 — when true, formatting
   * changes are still allowed even when edit protection is enforced.
   * Directly relevant to the tracked-changes workflow: combined with
   * `edit: 'trackedChanges'`, this forces all content edits to be tracked
   * while still permitting formatting adjustments.
   */
  formatting?: boolean;
  cryptProviderType?: string;
  cryptAlgorithmClass?: string;
  cryptAlgorithmType?: string;
  cryptAlgorithmSid?: number;
  cryptSpinCount?: number;
  hash?: string;
  salt?: string;
  /**
   * Modern crypto attributes per ISO/IEC 29500-4 §13 (Word 2013+):
   * - `algorithmName` names the strong hash algorithm (e.g. "SHA-512"),
   *   replacing the legacy `cryptAlgorithmSid` lookup-table reference.
   * - `hashValue` is the base64-encoded password hash.
   * - `saltValue` is the base64-encoded salt.
   *
   * These are emitted alongside / instead of the legacy `hash` / `salt`
   * / `cryptAlgorithmSid` attributes. Previously dropped on round-trip.
   */
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
}

/**
 * Revision view settings per ECMA-376 CT_TrackChangesView
 */
export interface RevisionViewSettings {
  showInsertionsAndDeletions: boolean;
  showFormatting: boolean;
  showInkAnnotations: boolean;
  /**
   * `w:markup` per CT_TrackChangesView §17.15.1.77 — when false, all
   * revision markup is hidden in the reviewer pane (no balloons, no
   * strikethrough/inserted styling). Defaults to true when absent.
   */
  showMarkup?: boolean;
  /**
   * `w:comments` per CT_TrackChangesView §17.15.1.77 — when false,
   * comment balloons are hidden. Defaults to true when absent.
   */
  showComments?: boolean;
}

/**
 * Track changes and related settings passed to DocumentGenerator.generateSettings()
 */
export interface TrackChangesSettings {
  trackChangesEnabled?: boolean;
  trackFormatting?: boolean;
  revisionView?: RevisionViewSettings;
  rsidRoot?: string;
  rsids?: string[];
  documentProtection?: DocumentProtection;
}

/**
 * Information about webSettings.xml per ECMA-376 CT_WebSettings
 */
export interface WebSettingsInfo {
  divCount: number;
  optimizeForBrowser: boolean;
  allowPNG: boolean;
  relyOnVML: boolean;
  doNotRelyOnCSS: boolean;
  doNotSaveAsSingleFile: boolean;
  doNotOrganizeInFolder: boolean;
  doNotUseLongFileNames: boolean;
  pixelsPerInch?: number;
  targetScreenSz?: string;
  encoding?: string;
}
