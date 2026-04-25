/**
 * Plugin extension point for custom validation rules. Built-in rules
 * live in `ValidationRules.ts`; this registry lets consumers add their
 * own without modifying framework code. Custom rules run alongside the
 * built-ins when `RevisionValidator.validate()` executes — the validator
 * appends `runAll()` issues to its `infos`/`warnings`/`errors` buckets
 * based on each rule's declared severity.
 */
import type { Document } from '../core/Document';
import { KeyedRegistry } from '../utils/KeyedRegistry';

/**
 * Severity for custom validation rules. Distinct from the built-in
 * `ValidationSeverity` enum so the public API surface has no name
 * collision.
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
  private readonly inner = new KeyedRegistry<CustomValidationRule>('ValidationRuleRegistry');

  register(rule: CustomValidationRule): void {
    this.inner.register(rule.code, rule);
  }

  unregister(code: string): boolean {
    return this.inner.unregister(code);
  }

  has(code: string): boolean {
    return this.inner.has(code);
  }

  get(code: string): CustomValidationRule | undefined {
    return this.inner.get(code);
  }

  getAll(): CustomValidationRule[] {
    return this.inner.values();
  }

  /** Run every registered rule and concatenate the issues. A throwing
   *  rule is captured as a synthetic 'error' issue so consumers see it. */
  runAll(doc: Document): CustomValidationIssue[] {
    const out: CustomValidationIssue[] = [];
    for (const rule of this.inner.values()) {
      try {
        out.push(...rule.validate(doc));
      } catch {
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
    this.inner.clear();
  }
}

export const ValidationRuleRegistry = new ValidationRuleRegistryImpl();
