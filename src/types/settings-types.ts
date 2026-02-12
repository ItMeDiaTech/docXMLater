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
  cryptProviderType?: string;
  cryptAlgorithmClass?: string;
  cryptAlgorithmType?: string;
  cryptAlgorithmSid?: number;
  cryptSpinCount?: number;
  hash?: string;
  salt?: string;
}

/**
 * Revision view settings per ECMA-376 CT_TrackChangesView
 */
export interface RevisionViewSettings {
  showInsertionsAndDeletions: boolean;
  showFormatting: boolean;
  showInkAnnotations: boolean;
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
