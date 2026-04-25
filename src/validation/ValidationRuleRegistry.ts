/**
 * Plugin extension point for custom validation rules.
 *
 * Existing built-in rules live in `ValidationRules.ts`. This registry
 * lets consumers add their own without modifying framework code. Custom
 * rules are evaluated alongside built-ins when `RevisionValidator`
 * runs validation passes.
 */
import type { Document } from '../core/Document';

/**
 * Severity for custom validation rules.
 * Distinct from `ValidationSeverity` (built-in rule severity enum) to
 * avoid name collision in the public API surface.
 */
export type CustomValidationSeverity = 'error' | 'warning' | 'info';

export interface CustomValidationIssue {
  ruleCode: string;
  severity: CustomValidationSeverity;
  message: string;
  /** Optional element identifier or path for navigation in tooling. */
  location?: string;
}

export interface CustomValidationRule {
  /** Unique rule identifier — must not collide with built-in rule codes. */
  code: string;
  /** Human-readable description of what the rule checks. */
  description: string;
  severity: CustomValidationSeverity;
  /**
   * Run the rule against a document and return any issues found.
   * Synchronous — async rules should be wrapped externally.
   */
  validate(doc: Document): CustomValidationIssue[];
}

class ValidationRuleRegistryImpl {
  private rules = new Map<string, CustomValidationRule>();

  register(rule: CustomValidationRule): void {
    if (this.rules.has(rule.code)) {
      throw new Error(`ValidationRuleRegistry: rule "${rule.code}" already registered`);
    }
    this.rules.set(rule.code, rule);
  }

  unregister(code: string): boolean {
    return this.rules.delete(code);
  }

  has(code: string): boolean {
    return this.rules.has(code);
  }

  get(code: string): CustomValidationRule | undefined {
    return this.rules.get(code);
  }

  getAll(): CustomValidationRule[] {
    return [...this.rules.values()];
  }

  /** Run every registered rule and concatenate the issues. */
  runAll(doc: Document): CustomValidationIssue[] {
    const out: CustomValidationIssue[] = [];
    for (const rule of this.rules.values()) {
      try {
        out.push(...rule.validate(doc));
      } catch {
        // A throwing rule must not abort other rules. Surface as a synthetic
        // issue so the consumer notices their rule misbehaved.
        out.push({
          ruleCode: rule.code,
          severity: 'error',
          message: `Rule "${rule.code}" threw during validation`,
        });
      }
    }
    return out;
  }

  clear(): void {
    this.rules.clear();
  }
}

export const ValidationRuleRegistry = new ValidationRuleRegistryImpl();
