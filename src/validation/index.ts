/**
 * Validation Module
 *
 * Provides validation and auto-fix capabilities for ECMA-376 revision compliance.
 *
 * @module validation
 */

// Rule definitions and types
export {
  REVISION_RULES,
  ValidationSeverity,
  ValidationIssue,
  ValidationRule,
  ValidationOptions,
  AutoFixOptions,
  ValidationResult,
  FixAction,
  AutoFixResult,
  createIssueFromRule,
  getRuleByCode,
  getRulesBySeverity,
  getAutoFixableRules,
} from './ValidationRules.js';

// Validator class
export { RevisionValidator } from './RevisionValidator.js';

// Auto-fixer class
export { RevisionAutoFixer } from './RevisionAutoFixer.js';
